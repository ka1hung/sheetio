// Package sheetio reads and writes xlsx and ods spreadsheets using only the Go standard library.
//
// Scope: exporting and importing tabular data (multiple sheets, strings, numbers, booleans, dates,
// basic styles, column widths, row heights, merged cells, freeze panes, auto filter). Formula evaluation, charts, pivot tables and images are not supported.
//
// The API accepts only io.Reader / io.ReaderAt / io.Writer and never touches file paths;
// opening and creating files is left to the caller.
package sheetio

import (
	"errors"
	"fmt"
	"math"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

// Format is a spreadsheet file format.
type Format int

const (
	FormatUnknown Format = iota
	FormatXLSX
	FormatODS
)

func (f Format) String() string {
	switch f {
	case FormatXLSX:
		return "xlsx"
	case FormatODS:
		return "ods"
	}
	return "unknown"
}

// Ext returns the file extension, including the dot.
func (f Format) Ext() string {
	switch f {
	case FormatXLSX:
		return ".xlsx"
	case FormatODS:
		return ".ods"
	}
	return ""
}

// ContentType returns the HTTP Content-Type.
func (f Format) ContentType() string {
	switch f {
	case FormatXLSX:
		return mimeXLSX
	case FormatODS:
		return mimeODS
	}
	return "application/octet-stream"
}

const (
	mimeXLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	mimeODS  = "application/vnd.oasis.opendocument.spreadsheet"
)

// Limits defined by the Excel specification.
const (
	MaxSheetRows  = 1048576
	MaxSheetCols  = 16384
	MaxCellChars  = 32767
	MaxSheetName  = 31
	invalidSheetC = `[]:*?/\`
)

// CellKind is the type of a cell value.
type CellKind uint8

const (
	KindEmpty CellKind = iota
	KindString
	KindNumber
	KindBool
	KindDate
)

// Cell is a single cell.
type Cell struct {
	Kind  CellKind
	Str   string
	Num   float64
	Bool  bool
	Time  time.Time // used by KindDate; the wall-clock time is written, without time zone conversion
	Style int       // 0 = default; other values come from Workbook.AddStyle
}

// Str creates a string cell.
func Str(s string) Cell { return Cell{Kind: KindString, Str: s} }

// Num creates a number cell.
func Num(f float64) Cell { return Cell{Kind: KindNumber, Num: f} }

// Int creates an integer cell. Integers above 2^53 lose precision; use Str instead.
func Int(i int64) Cell { return Cell{Kind: KindNumber, Num: float64(i)} }

// Bool creates a boolean cell.
func Bool(b bool) Cell { return Cell{Kind: KindBool, Bool: b} }

// Date creates a date/time cell.
func Date(t time.Time) Cell { return Cell{Kind: KindDate, Time: t} }

// Empty creates a blank cell (combine with WithStyle for a styled empty cell).
func Empty() Cell { return Cell{} }

// WithStyle returns the cell with the given style applied.
func (c Cell) WithStyle(id int) Cell { c.Style = id; return c }

// IsEmpty reports whether the cell is blank.
func (c Cell) IsEmpty() bool { return c.Kind == KindEmpty }

// String returns the cell value as a string, suitable for display and import matching.
func (c Cell) String() string {
	switch c.Kind {
	case KindString:
		return c.Str
	case KindNumber:
		return strconv.FormatFloat(c.Num, 'f', -1, 64)
	case KindBool:
		if c.Bool {
			return "TRUE"
		}
		return "FALSE"
	case KindDate:
		if hasClock(c.Time) {
			return c.Time.Format("2006-01-02 15:04:05")
		}
		return c.Time.Format("2006-01-02")
	}
	return ""
}

func hasClock(t time.Time) bool {
	h, m, s := t.Clock()
	return h != 0 || m != 0 || s != 0 || t.Nanosecond() != 0
}

// NumFmt is a built-in number format.
type NumFmt uint8

const (
	NumFmtGeneral   NumFmt = iota
	NumFmtInteger          // 0
	NumFmtDecimal2         // 0.00
	NumFmtThousands        // #,##0
	NumFmtDate             // yyyy-mm-dd
	NumFmtDateTime         // yyyy-mm-dd hh:mm:ss
)

func (f NumFmt) isDate() bool { return f == NumFmtDate || f == NumFmtDateTime }

// HAlign is a horizontal alignment.
type HAlign uint8

const (
	AlignDefault HAlign = iota
	AlignLeft
	AlignCenter
	AlignRight
)

// Style is a cell style (deliberately limited to what reports commonly need).
type Style struct {
	Bold      bool
	FillColor string // background color "RRGGBB"; empty for none
	Border    bool   // thin border on all four sides
	Align     HAlign
	Wrap      bool // wrap text
	NumFmt    NumFmt
}

// Range is a cell range (merges, auto filter), 0-based and inclusive.
type Range struct{ Row1, Col1, Row2, Col2 int }

// Sheet is a worksheet.
type Sheet struct {
	Name       string
	ColWidths  []float64       // column widths in characters; 0 means default
	RowHeights map[int]float64 // row heights in points (1-409), keyed by 0-based row; applies only to rows that are written
	Rows       [][]Cell
	Merges     []Range
	FreezeRows int    // number of top rows to freeze (0 = none)
	FreezeCols int    // number of left columns to freeze (0 = none)
	AutoFilter *Range // auto filter range; Row2 is extended to the last data row when written
}

// MaxRowHeight is the maximum row height in points, same as Excel.
const MaxRowHeight = 409

// AddRow appends a row.
func (s *Sheet) AddRow(cells ...Cell) { s.Rows = append(s.Rows, cells) }

// SetRowHeight sets a row height in points (1-409; the Excel default is about 15).
func (s *Sheet) SetRowHeight(row int, height float64) {
	if s.RowHeights == nil {
		s.RowHeights = map[int]float64{}
	}
	s.RowHeights[row] = height
}

// Freeze freezes the top rows and the left cols (e.g. Freeze(1, 0) keeps the header row visible).
func (s *Sheet) Freeze(rows, cols int) { s.FreezeRows, s.FreezeCols = rows, cols }

// SetAutoFilter adds filter buttons to columns firstCol..lastCol of headerRow (0-based).
// The filter range extends to the last data row.
func (s *Sheet) SetAutoFilter(headerRow, firstCol, lastCol int) {
	s.AutoFilter = &Range{headerRow, firstCol, headerRow, lastCol}
}

// filterRange returns the effective filter range: Row2 extends at least to the last row (rows is the number of rows written).
func (s *Sheet) filterRange(rows int) (Range, bool) {
	if s.AutoFilter == nil {
		return Range{}, false
	}
	r := *s.AutoFilter
	if rows-1 > r.Row2 {
		r.Row2 = rows - 1
	}
	return r, true
}

// SetColWidth sets a column width in characters.
func (s *Sheet) SetColWidth(col int, width float64) {
	for len(s.ColWidths) <= col {
		s.ColWidths = append(s.ColWidths, 0)
	}
	s.ColWidths[col] = width
}

// Merge merges a range of cells (0-based, inclusive).
func (s *Sheet) Merge(row1, col1, row2, col2 int) {
	s.Merges = append(s.Merges, Range{row1, col1, row2, col2})
}

// Cell returns the cell at row, col, or a blank cell when out of range.
func (s *Sheet) Cell(row, col int) Cell {
	if row < 0 || row >= len(s.Rows) || col < 0 || col >= len(s.Rows[row]) {
		return Cell{}
	}
	return s.Rows[row][col]
}

// Strings returns the whole sheet as a 2D string slice (the most common way to import).
func (s *Sheet) Strings() [][]string {
	out := make([][]string, len(s.Rows))
	for i, row := range s.Rows {
		out[i] = make([]string, len(row))
		for j, c := range row {
			out[i][j] = c.String()
		}
	}
	return out
}

// Workbook is a spreadsheet document.
type Workbook struct {
	Sheets []*Sheet
	Styles []Style
}

// New creates an empty workbook.
func New() *Workbook { return &Workbook{} }

// AddSheet adds a worksheet.
func (wb *Workbook) AddSheet(name string) *Sheet {
	s := &Sheet{Name: name}
	wb.Sheets = append(wb.Sheets, s)
	return s
}

// AddStyle registers a style and returns its ID (starting from 1).
func (wb *Workbook) AddStyle(st Style) int {
	wb.Styles = append(wb.Styles, st)
	return len(wb.Styles)
}

// Sheet returns the sheet with the given name (case-insensitive), or nil if not found.
func (wb *Workbook) Sheet(name string) *Sheet {
	for _, s := range wb.Sheets {
		if strings.EqualFold(s.Name, name) {
			return s
		}
	}
	return nil
}

func (wb *Workbook) style(id int) Style {
	if id <= 0 || id > len(wb.Styles) {
		return Style{}
	}
	return wb.Styles[id-1]
}

// ErrInvalid means the workbook content is not valid and cannot be written.
var ErrInvalid = errors.New("sheetio: invalid workbook")

// Representable date range (the xlsx 1900 date system has a leap-year bug before 1900-03-01, so earlier dates are not supported).
var (
	minDate = time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC)
	maxDate = time.Date(9999, 12, 31, 23, 59, 59, 0, time.UTC)
)

// ValidateSheetName checks whether a sheet name follows the Excel rules.
func ValidateSheetName(name string) error {
	n := utf8.RuneCountInString(name)
	if n == 0 || n > MaxSheetName {
		return fmt.Errorf("%w: sheet name %q must be 1-%d characters", ErrInvalid, name, MaxSheetName)
	}
	if strings.ContainsAny(name, invalidSheetC) {
		return fmt.Errorf("%w: sheet name %q contains one of %s", ErrInvalid, name, invalidSheetC)
	}
	if strings.HasPrefix(name, "'") || strings.HasSuffix(name, "'") {
		return fmt.Errorf("%w: sheet name %q cannot start or end with '", ErrInvalid, name)
	}
	if !utf8.ValidString(name) {
		return fmt.Errorf("%w: sheet name is not valid UTF-8", ErrInvalid)
	}
	return nil
}

// Validate checks whether the workbook can be written. Write calls it automatically.
func (wb *Workbook) Validate() error {
	if wb == nil || len(wb.Sheets) == 0 {
		return fmt.Errorf("%w: workbook has no sheets", ErrInvalid)
	}
	seen := map[string]bool{}
	for _, st := range wb.Styles {
		if st.FillColor != "" && !isHexColor(st.FillColor) {
			return fmt.Errorf("%w: fill color %q must be RRGGBB", ErrInvalid, st.FillColor)
		}
		if st.NumFmt > NumFmtDateTime || st.Align > AlignRight {
			return fmt.Errorf("%w: unknown number format or alignment", ErrInvalid)
		}
	}
	for _, sh := range wb.Sheets {
		if sh == nil {
			return fmt.Errorf("%w: nil sheet", ErrInvalid)
		}
		if err := ValidateSheetName(sh.Name); err != nil {
			return err
		}
		key := strings.ToLower(sh.Name)
		if seen[key] {
			return fmt.Errorf("%w: duplicate sheet name %q", ErrInvalid, sh.Name)
		}
		seen[key] = true
		if len(sh.Rows) > MaxSheetRows {
			return fmt.Errorf("%w: sheet %q has more than %d rows", ErrInvalid, sh.Name, MaxSheetRows)
		}
		if len(sh.ColWidths) > MaxSheetCols {
			return fmt.Errorf("%w: sheet %q has more than %d column widths", ErrInvalid, sh.Name, MaxSheetCols)
		}
		for _, w := range sh.ColWidths {
			if w < 0 || w > 255 || math.IsNaN(w) {
				return fmt.Errorf("%w: sheet %q column width must be 0-255", ErrInvalid, sh.Name)
			}
		}
		for r, h := range sh.RowHeights {
			if r < 0 || r >= MaxSheetRows || !(h >= 1 && h <= MaxRowHeight) {
				return fmt.Errorf("%w: sheet %q row %d height must be 1-%d points", ErrInvalid, sh.Name, r+1, MaxRowHeight)
			}
		}
		if sh.FreezeRows < 0 || sh.FreezeRows >= MaxSheetRows || sh.FreezeCols < 0 || sh.FreezeCols >= MaxSheetCols {
			return fmt.Errorf("%w: sheet %q invalid freeze %d rows / %d cols", ErrInvalid, sh.Name, sh.FreezeRows, sh.FreezeCols)
		}
		if f := sh.AutoFilter; f != nil {
			if f.Row1 < 0 || f.Col1 < 0 || f.Row2 < f.Row1 || f.Col2 < f.Col1 || f.Row2 >= MaxSheetRows || f.Col2 >= MaxSheetCols {
				return fmt.Errorf("%w: sheet %q invalid auto filter range %+v", ErrInvalid, sh.Name, *f)
			}
		}
		for r, row := range sh.Rows {
			if len(row) > MaxSheetCols {
				return fmt.Errorf("%w: sheet %q row %d has more than %d columns", ErrInvalid, sh.Name, r+1, MaxSheetCols)
			}
			for c, cell := range row {
				if err := wb.validateCell(cell); err != nil {
					return fmt.Errorf("%w: sheet %q cell %s: %v", ErrInvalid, sh.Name, CellRef(r, c), err)
				}
			}
		}
		for _, m := range sh.Merges {
			if m.Row1 < 0 || m.Col1 < 0 || m.Row2 < m.Row1 || m.Col2 < m.Col1 ||
				m.Row2 >= MaxSheetRows || m.Col2 >= MaxSheetCols || (m.Row1 == m.Row2 && m.Col1 == m.Col2) {
				return fmt.Errorf("%w: sheet %q invalid merge range %+v", ErrInvalid, sh.Name, m)
			}
		}
		if err := checkMergeOverlap(sh.Merges); err != nil {
			return fmt.Errorf("%w: sheet %q: %v", ErrInvalid, sh.Name, err)
		}
	}
	return nil
}

func (wb *Workbook) validateCell(c Cell) error {
	if c.Style < 0 || c.Style > len(wb.Styles) {
		return fmt.Errorf("unknown style id %d", c.Style)
	}
	switch c.Kind {
	case KindEmpty, KindBool:
	case KindString:
		if !utf8.ValidString(c.Str) {
			return errors.New("string is not valid UTF-8")
		}
		if utf8.RuneCountInString(c.Str) > MaxCellChars {
			return fmt.Errorf("string longer than %d characters", MaxCellChars)
		}
	case KindNumber:
		if math.IsNaN(c.Num) || math.IsInf(c.Num, 0) {
			return errors.New("number is NaN or Inf")
		}
	case KindDate:
		w := wallUTC(c.Time)
		if w.Before(minDate) || w.After(maxDate) {
			return errors.New("date must be between 1900-03-01 and 9999-12-31")
		}
	default:
		return fmt.Errorf("unknown cell kind %d", c.Kind)
	}
	return nil
}

func checkMergeOverlap(ms []Range) error {
	// there are usually few merges, so O(n^2) is fine; the cap keeps large inputs fast.
	if len(ms) > 10000 {
		return errors.New("too many merged ranges (max 10000)")
	}
	for i := range ms {
		for j := i + 1; j < len(ms); j++ {
			a, b := ms[i], ms[j]
			if a.Row1 <= b.Row2 && b.Row1 <= a.Row2 && a.Col1 <= b.Col2 && b.Col1 <= a.Col2 {
				return fmt.Errorf("merged ranges %s and %s overlap", a.ref(), b.ref())
			}
		}
	}
	return nil
}

func (r Range) ref() string { return CellRef(r.Row1, r.Col1) + ":" + CellRef(r.Row2, r.Col2) }

func isHexColor(s string) bool {
	if len(s) != 6 {
		return false
	}
	for i := 0; i < 6; i++ {
		c := s[i]
		if !(c >= '0' && c <= '9' || c >= 'a' && c <= 'f' || c >= 'A' && c <= 'F') {
			return false
		}
	}
	return true
}

// ColName converts a 0-based column index to A, B, ... AA.
func ColName(col int) string {
	var buf [4]byte
	i := len(buf)
	col++
	for col > 0 {
		col--
		i--
		buf[i] = byte('A' + col%26)
		col /= 26
	}
	return string(buf[i:])
}

// CellRef converts a 0-based row and column to an A1 reference.
func CellRef(row, col int) string { return ColName(col) + strconv.Itoa(row+1) }

// parseCellRef parses an A1 reference and returns the 0-based row and column.
func parseCellRef(ref string) (row, col int, ok bool) {
	i := 0
	for i < len(ref) && ref[i] >= 'A' && ref[i] <= 'Z' {
		if i >= 3 {
			return 0, 0, false
		}
		col = col*26 + int(ref[i]-'A'+1)
		i++
	}
	if i == 0 || i == len(ref) || len(ref)-i > 7 {
		return 0, 0, false
	}
	for j := i; j < len(ref); j++ {
		if ref[j] < '0' || ref[j] > '9' {
			return 0, 0, false
		}
		row = row*10 + int(ref[j]-'0')
	}
	if row == 0 {
		return 0, 0, false
	}
	return row - 1, col - 1, true
}

// wallUTC moves the wall-clock time (date and clock) to UTC so time zones cannot shift the date.
func wallUTC(t time.Time) time.Time {
	y, mo, d := t.Date()
	h, mi, s := t.Clock()
	return time.Date(y, mo, d, h, mi, s, t.Nanosecond(), time.UTC)
}
