package sheetio

import (
	"errors"
	"fmt"
	"io"
)

// SheetDef defines a sheet for streaming output. Column widths and merges must be given up front
// (xlsx column definitions come before the rows; ods needs merges while the rows are written).
type SheetDef struct {
	Name       string
	ColWidths  []float64
	RowHeights map[int]float64 // row heights in points, keyed by 0-based row
	Merges     []Range
	FreezeRows int    // number of top rows to freeze
	FreezeCols int    // number of left columns to freeze
	AutoFilter *Range // auto filter; Row2 extends to the last written row, so the header row alone is usually enough
}

// ErrClosed means the StreamWriter has already been closed.
var ErrClosed = errors.New("sheetio: stream writer closed")

// StreamWriter writes a spreadsheet row by row; memory use does not depend on the number of rows:
//   - xlsx: each row is compressed straight into the output, nothing is buffered.
//   - ods: only the compressed content.xml is kept in memory (usually under 1/10 of the raw XML),
//     because Excel requires the CRC and sizes in the zip local headers.
//
// No temporary files are used, so no writable temp directory is needed.
//
// Usage:
//
//	sw, _ := sheetio.NewStreamWriter(w, sheetio.FormatXLSX, styles, []sheetio.SheetDef{{Name: "Data"}})
//	for rows.Next() { sw.WriteRow(...) }
//	sw.NextSheet() // when there is a second sheet
//	err := sw.Close()
type StreamWriter struct {
	wb     *Workbook // only Styles and the sheet definitions are used (no rows)
	st     *styleTable
	sink   sheetSink
	cur    int
	row    int
	err    error // write error (not recoverable)
	closed bool
	check  bool
}

type sheetSink interface {
	begin(i int) error
	row(i, r int, cells []Cell) error
	end(i, rows int) error
	close() error
}

// NewStreamWriter creates a streaming writer. Style IDs follow the same rule as Workbook.AddStyle
// (styles[0] is ID 1).
func NewStreamWriter(w io.Writer, f Format, styles []Style, sheets []SheetDef) (*StreamWriter, error) {
	wb := &Workbook{Styles: styles}
	for _, d := range sheets {
		wb.Sheets = append(wb.Sheets, &Sheet{
			Name: d.Name, ColWidths: d.ColWidths, RowHeights: d.RowHeights, Merges: d.Merges,
			FreezeRows: d.FreezeRows, FreezeCols: d.FreezeCols, AutoFilter: d.AutoFilter,
		})
	}
	if err := wb.Validate(); err != nil {
		return nil, err
	}
	return newStreamWriter(w, f, wb, true)
}

func newStreamWriter(w io.Writer, f Format, wb *Workbook, check bool) (*StreamWriter, error) {
	sw := &StreamWriter{wb: wb, st: newStyleTable(wb), cur: -1, check: check}
	var err error
	switch f {
	case FormatXLSX:
		sw.sink, err = newXLSXSink(w, wb, sw.st)
	case FormatODS:
		sw.sink, err = newODSSink(w, wb, sw.st)
	default:
		return nil, fmt.Errorf("%w: %v", ErrUnsupported, f)
	}
	if err != nil {
		return nil, err
	}
	return sw, nil
}

func (sw *StreamWriter) fail(err error) error {
	if err != nil && sw.err == nil {
		sw.err = err
	}
	return err
}

// NextSheet finishes the current sheet and starts the next one (in SheetDef order).
// It is not needed before the first WriteRow, which starts the first sheet automatically.
func (sw *StreamWriter) NextSheet() error {
	if sw.closed {
		return ErrClosed
	}
	if sw.err != nil {
		return sw.err
	}
	if sw.cur+1 >= len(sw.wb.Sheets) {
		return fmt.Errorf("%w: only %d sheets defined", ErrInvalid, len(sw.wb.Sheets))
	}
	if sw.cur >= 0 {
		if err := sw.sink.end(sw.cur, sw.row); err != nil {
			return sw.fail(err)
		}
	}
	sw.cur++
	sw.row = 0
	return sw.fail(sw.sink.begin(sw.cur))
}

// WriteRow writes the next row. Invalid data returns ErrInvalid; that row is not written and writing can continue.
// cells can be reused as soon as the call returns.
func (sw *StreamWriter) WriteRow(cells ...Cell) error { return sw.writeRow(cells) }

func (sw *StreamWriter) writeRow(cells []Cell) error {
	if sw.closed {
		return ErrClosed
	}
	if sw.err != nil {
		return sw.err
	}
	if sw.cur < 0 {
		if err := sw.NextSheet(); err != nil {
			return err
		}
	}
	name := sw.wb.Sheets[sw.cur].Name
	if sw.row >= MaxSheetRows {
		return fmt.Errorf("%w: sheet %q exceeds %d rows", ErrInvalid, name, MaxSheetRows)
	}
	if len(cells) > MaxSheetCols {
		return fmt.Errorf("%w: sheet %q row %d has more than %d columns", ErrInvalid, name, sw.row+1, MaxSheetCols)
	}
	if sw.check {
		for c, cell := range cells {
			if err := sw.wb.validateCell(cell); err != nil {
				return fmt.Errorf("%w: sheet %q cell %s: %v", ErrInvalid, name, CellRef(sw.row, c), err)
			}
		}
	}
	if err := sw.sink.row(sw.cur, sw.row, cells); err != nil {
		return sw.fail(err)
	}
	sw.row++
	return nil
}

// Close finishes all sheets (sheets that were never reached are left empty) and completes the file. It must be called.
func (sw *StreamWriter) Close() error {
	if sw.closed {
		return sw.err
	}
	if sw.err != nil {
		sw.closed = true
		return sw.err
	}
	for sw.cur < len(sw.wb.Sheets)-1 {
		if err := sw.NextSheet(); err != nil {
			sw.closed = true
			return err
		}
	}
	sw.closed = true
	if err := sw.sink.end(sw.cur, sw.row); err != nil {
		return sw.fail(err)
	}
	return sw.fail(sw.sink.close())
}
