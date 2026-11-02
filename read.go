package sheetio

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"strings"
	"time"
)

// ErrLimit means the input exceeds a configured Limits value (the file is too large or malformed).
var ErrLimit = errors.New("sheetio: limit exceeded")

// ErrUnsupported means the file format is not supported.
var ErrUnsupported = errors.New("sheetio: unsupported file format")

// ErrCorrupt means the file is damaged or does not follow the format.
var ErrCorrupt = errors.New("sheetio: corrupt file")

// Limits caps the resources used while reading (file size, decompressed size, rows, cells, ...).
// Zero fields use the value from DefaultLimits.
type Limits struct {
	MaxFileSize  int64 // size of the (compressed) file; default 50 MB
	MaxEntries   int   // number of zip entries; default 10000
	MaxEntrySize int64 // decompressed size of one zip entry; default 256 MB
	MaxTotalSize int64 // total decompressed size of all entries read; default 512 MB
	MaxSheets    int   // number of sheets; default 256
	MaxRows      int   // rows per sheet; default 1048576
	MaxCols      int   // columns per sheet; default 16384
	MaxCells     int   // total cells plus row slots allocated for the workbook (bounds memory); default 5,000,000
	MaxCellChars int   // characters per cell; default 32767
}

// DefaultLimits returns the default limits.
func DefaultLimits() Limits {
	return Limits{
		MaxFileSize:  50 << 20,
		MaxEntries:   10000,
		MaxEntrySize: 256 << 20,
		MaxTotalSize: 512 << 20,
		MaxSheets:    256,
		MaxRows:      MaxSheetRows,
		MaxCols:      MaxSheetCols,
		MaxCells:     5_000_000,
		MaxCellChars: MaxCellChars,
	}
}

func (l Limits) withDefaults() Limits {
	d := DefaultLimits()
	if l.MaxFileSize <= 0 {
		l.MaxFileSize = d.MaxFileSize
	}
	if l.MaxEntries <= 0 {
		l.MaxEntries = d.MaxEntries
	}
	if l.MaxEntrySize <= 0 {
		l.MaxEntrySize = d.MaxEntrySize
	}
	if l.MaxTotalSize <= 0 {
		l.MaxTotalSize = d.MaxTotalSize
	}
	if l.MaxSheets <= 0 {
		l.MaxSheets = d.MaxSheets
	}
	if l.MaxRows <= 0 || l.MaxRows > MaxSheetRows {
		l.MaxRows = d.MaxRows
	}
	if l.MaxCols <= 0 || l.MaxCols > MaxSheetCols {
		l.MaxCols = d.MaxCols
	}
	if l.MaxCells <= 0 {
		l.MaxCells = d.MaxCells
	}
	if l.MaxCellChars <= 0 {
		l.MaxCellChars = d.MaxCellChars
	}
	return l
}

// ReadOptions are reading options. A nil *ReadOptions uses the defaults.
type ReadOptions struct {
	Limits Limits
	// Location is the time zone dates are placed in (dates in the file have no time zone); default time.Local.
	Location *time.Location
}

func (o *ReadOptions) normalize() ReadOptions {
	var n ReadOptions
	if o != nil {
		n = *o
	}
	n.Limits = n.Limits.withDefaults()
	if n.Location == nil {
		n.Location = time.Local
	}
	return n
}

// Read loads the whole file into memory (xlsx or ods, detected from the content, not the file name).
// multipart.File, *os.File and *bytes.Reader all implement io.ReaderAt.
// For large files use ReadRows to process rows one at a time.
func Read(r io.ReaderAt, size int64, opt *ReadOptions) (wb *Workbook, f Format, err error) {
	o := opt.normalize()
	// Any unexpected panic is turned into an error instead of crashing the caller.
	defer func() {
		if p := recover(); p != nil {
			wb, f, err = nil, FormatUnknown, fmt.Errorf("%w: internal error: %v", ErrCorrupt, p)
		}
	}()
	return read(r, size, &o)
}

func read(r io.ReaderAt, size int64, o *ReadOptions) (*Workbook, Format, error) {
	rd := &reader{lim: &o.Limits, loc: o.Location}
	c := &collector{rd: rd, wb: New()}
	f, err := readInto(r, size, rd, c)
	if err != nil {
		return nil, f, err
	}
	return c.wb, f, nil
}

// StreamDefaultMaxCells is the MaxCells limit ReadRows uses when none is set (100 million).
const StreamDefaultMaxCells = 100_000_000

// ErrSkipSheet is returned by a RowFunc to skip the rest of the current sheet and continue with the next one.
var ErrSkipSheet = errors.New("sheetio: skip rest of sheet")

// RowFunc is called once for each non-empty row.
// row is the 0-based row number (empty rows are skipped, so numbers can have gaps);
// trailing empty cells are already removed, and cells is reused after the callback returns, so copy anything you keep.
// Return ErrSkipSheet to skip the rest of the sheet; any other error stops reading and is returned by ReadRows as is.
type RowFunc func(sheet string, row int, cells []Cell) error

// ReadRows reads the file as a stream, passing each row to fn without keeping the sheet; memory use does not depend on the number of rows
// (the xlsx shared string table is still loaded in full).
// No temporary files are used. Here Limits.MaxCells limits the total number of cells passed to fn.
// When not set, the defaults are wider than for Read: MaxCells 100 million, 2 GB per entry, 4 GB in total;
// everything else is the same as DefaultLimits. Set smaller limits for user uploads.
func ReadRows(r io.ReaderAt, size int64, opt *ReadOptions, fn RowFunc) (f Format, err error) {
	// memory use is flat in streaming mode, so the default limits are wider (only CPU time matters);
	// values set by the caller are used as is.
	var o0 ReadOptions
	if opt != nil {
		o0 = *opt
	}
	if o0.Limits.MaxCells <= 0 {
		o0.Limits.MaxCells = StreamDefaultMaxCells
	}
	if o0.Limits.MaxEntrySize <= 0 {
		o0.Limits.MaxEntrySize = 2 << 30
	}
	if o0.Limits.MaxTotalSize <= 0 {
		o0.Limits.MaxTotalSize = 4 << 30
	}
	o := o0.normalize()
	defer func() {
		if p := recover(); p != nil {
			f, err = FormatUnknown, fmt.Errorf("%w: internal error: %v", ErrCorrupt, p)
		}
	}()
	rd := &reader{lim: &o.Limits, loc: o.Location}
	return readInto(r, size, rd, funcHandler(fn))
}

func readInto(r io.ReaderAt, size int64, rd *reader, h rowHandler) (Format, error) {
	a, err := openZip(r, size, rd.lim)
	if err != nil {
		return FormatUnknown, err
	}
	f, err := a.detect()
	if err != nil {
		return FormatUnknown, err
	}
	switch f {
	case FormatXLSX:
		err = readXLSX(a, rd, h)
	case FormatODS:
		err = readODS(a, rd, h)
	}
	return f, err
}

// Detect returns the format without parsing the content.
func Detect(r io.ReaderAt, size int64) (Format, error) {
	lim := DefaultLimits()
	a, err := openZip(r, size, &lim)
	if err != nil {
		return FormatUnknown, err
	}
	return a.detect()
}

// ---- zip access with limits ----

type zipArchive struct {
	files map[string]*zip.File
	lim   *Limits
	total int64
}

var oleMagic = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

func openZip(r io.ReaderAt, size int64, lim *Limits) (*zipArchive, error) {
	if size <= 0 {
		return nil, fmt.Errorf("%w: empty file", ErrUnsupported)
	}
	if size > lim.MaxFileSize {
		return nil, fmt.Errorf("%w: file size %d > %d", ErrLimit, size, lim.MaxFileSize)
	}
	head := make([]byte, 8)
	if n, _ := r.ReadAt(head, 0); n == 8 && bytes.Equal(head, oleMagic) {
		return nil, fmt.Errorf("%w: legacy .xls or password-protected file", ErrUnsupported)
	}
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("%w: not a zip-based spreadsheet: %v", ErrUnsupported, err)
	}
	if len(zr.File) > lim.MaxEntries {
		return nil, fmt.Errorf("%w: %d zip entries > %d", ErrLimit, len(zr.File), lim.MaxEntries)
	}
	a := &zipArchive{files: make(map[string]*zip.File, len(zr.File)), lim: lim}
	for _, f := range zr.File {
		if _, dup := a.files[f.Name]; !dup {
			a.files[f.Name] = f
		}
	}
	return a, nil
}

func (a *zipArchive) has(name string) bool { _, ok := a.files[name]; return ok }

func (a *zipArchive) detect() (Format, error) {
	if f, ok := a.files["mimetype"]; ok && f.UncompressedSize64 < 128 {
		b, err := a.readAll("mimetype")
		if err == nil && strings.TrimSpace(string(b)) == mimeODS {
			if a.has("content.xml") {
				return FormatODS, nil
			}
			return FormatUnknown, fmt.Errorf("%w: ods without content.xml", ErrCorrupt)
		}
	}
	if a.has("[Content_Types].xml") {
		return FormatXLSX, nil
	}
	return FormatUnknown, fmt.Errorf("%w: neither xlsx nor ods", ErrUnsupported)
}

// open opens a zip entry and returns a reader that counts bytes and enforces the limits.
func (a *zipArchive) open(name string) (io.ReadCloser, error) {
	f, ok := a.files[name]
	if !ok {
		return nil, fmt.Errorf("%w: missing %s", ErrCorrupt, name)
	}
	if f.UncompressedSize64 > uint64(a.lim.MaxEntrySize) {
		return nil, fmt.Errorf("%w: %s declares %d bytes > %d", ErrLimit, name, f.UncompressedSize64, a.lim.MaxEntrySize)
	}
	rc, err := f.Open()
	if err != nil {
		return nil, fmt.Errorf("%w: %s: %v", ErrCorrupt, name, err)
	}
	return &countingReader{rc: rc, a: a, name: name}, nil
}

func (a *zipArchive) readAll(name string) ([]byte, error) {
	rc, err := a.open(name)
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

type countingReader struct {
	rc   io.ReadCloser
	a    *zipArchive
	name string
	n    int64
}

func (c *countingReader) Read(p []byte) (int, error) {
	n, err := c.rc.Read(p)
	c.n += int64(n)
	c.a.total += int64(n)
	if c.n > c.a.lim.MaxEntrySize {
		return n, fmt.Errorf("%w: %s exceeds %d bytes uncompressed", ErrLimit, c.name, c.a.lim.MaxEntrySize)
	}
	if c.a.total > c.a.lim.MaxTotalSize {
		return n, fmt.Errorf("%w: total uncompressed size exceeds %d bytes", ErrLimit, c.a.lim.MaxTotalSize)
	}
	return n, err
}

func (c *countingReader) Close() error { return c.rc.Close() }

// ---- row delivery ----

type rowHandler interface {
	startSheet(name string) error
	row(sheet string, r int, cells []Cell) error
	endSheet() error
}

type funcHandler RowFunc

func (funcHandler) startSheet(string) error                    { return nil }
func (f funcHandler) row(sh string, r int, cells []Cell) error { return f(sh, r, cells) }
func (funcHandler) endSheet() error                            { return nil }

type reader struct {
	lim    *Limits
	loc    *time.Location
	alloc  int // cells delivered + empty rows padded by the collector
	sheets int
}

func (rd *reader) charge(n int) error {
	rd.alloc += n
	if rd.alloc > rd.lim.MaxCells {
		return fmt.Errorf("%w: more than %d cells", ErrLimit, rd.lim.MaxCells)
	}
	return nil
}

func (rd *reader) startSheet(h rowHandler, name string) error {
	rd.sheets++
	if rd.sheets > rd.lim.MaxSheets {
		return fmt.Errorf("%w: more than %d sheets", ErrLimit, rd.lim.MaxSheets)
	}
	return h.startSheet(name)
}

// emit trims trailing empty cells and delivers a row; empty rows are not delivered.
func (rd *reader) emit(h rowHandler, sheet string, r int, cells []Cell) error {
	n := len(cells)
	for n > 0 && cells[n-1].IsEmpty() {
		n--
	}
	if n == 0 {
		return nil
	}
	if r < 0 || r >= rd.lim.MaxRows {
		return fmt.Errorf("%w: sheet %q row %d exceeds %d", ErrLimit, sheet, r+1, rd.lim.MaxRows)
	}
	if err := rd.charge(n); err != nil {
		return err
	}
	return h.row(sheet, r, cells[:n])
}

// collector gathers streamed rows into a Workbook (used by Read).
type collector struct {
	rd *reader
	wb *Workbook
	sh *Sheet
}

func (c *collector) startSheet(name string) error {
	c.sh = c.wb.AddSheet(name)
	return nil
}

func (c *collector) row(_ string, r int, cells []Cell) error {
	if r >= len(c.sh.Rows) {
		if err := c.rd.charge(r + 1 - len(c.sh.Rows)); err != nil {
			return err
		}
		for len(c.sh.Rows) <= r {
			c.sh.Rows = append(c.sh.Rows, nil)
		}
	}
	row := make([]Cell, len(cells))
	copy(row, cells)
	c.sh.Rows[r] = row
	return nil
}

func (c *collector) endSheet() error {
	trim(c.sh)
	return nil
}

// trim removes trailing empty cells from each row and trailing empty rows.
func trim(sh *Sheet) {
	for i, r := range sh.Rows {
		n := len(r)
		for n > 0 && r[n-1].IsEmpty() {
			n--
		}
		if n == 0 {
			sh.Rows[i] = nil
		} else {
			sh.Rows[i] = r[:n:n]
		}
	}
	n := len(sh.Rows)
	for n > 0 && len(sh.Rows[n-1]) == 0 {
		n--
	}
	sh.Rows = sh.Rows[:n]
}

// rawSkip skips the current element. Large data is parsed with RawToken (no namespace handling, about 15% faster),
// and Decoder.Skip only works with Token, so the depth is tracked here.
func rawSkip(d *xml.Decoder) error {
	for depth := 1; depth > 0; {
		tok, err := d.RawToken()
		if err != nil {
			if err == io.EOF {
				return io.ErrUnexpectedEOF
			}
			return err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

// nsResolver maps RawToken prefixes to namespace URIs (ODS prefixes are fixed by convention, but not guaranteed by the spec).
// Scoping is not tracked: redeclaring a prefix to a different URI inside a file is very rare, so the last declaration wins.
type nsResolver map[string]string

func (m nsResolver) start(t *xml.StartElement) {
	for _, a := range t.Attr {
		switch {
		case a.Name.Space == "xmlns":
			m[a.Name.Local] = a.Value
		case a.Name.Space == "" && a.Name.Local == "xmlns":
			m[""] = a.Value
		}
	}
	t.Name.Space = m[t.Name.Space]
	for i := range t.Attr {
		if sp := t.Attr[i].Name.Space; sp != "" && sp != "xmlns" {
			t.Attr[i].Name.Space = m[sp]
		}
	}
}

func (m nsResolver) end(t *xml.EndElement) { t.Name.Space = m[t.Name.Space] }

// textBuf is a string builder with a length limit.
type textBuf struct {
	b   strings.Builder
	max int // byte limit (character limit x 4)
}

func (t *textBuf) reset(maxChars int) { t.b.Reset(); t.max = maxChars * 4 }

func (t *textBuf) write(s string) error {
	if t.b.Len()+len(s) > t.max {
		return fmt.Errorf("%w: cell text longer than %d bytes", ErrLimit, t.max)
	}
	t.b.WriteString(s)
	return nil
}

func (t *textBuf) String() string { return t.b.String() }
