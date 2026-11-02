package sheetio

import (
	"bufio"
	"encoding/xml"
	"io"
	"strconv"
	"time"
)

// Write writes the whole workbook in the given format. The workbook is validated first.
// With Gin you can write directly to the response: Write(c.Writer, wb, sheetio.FormatXLSX).
// For large data use NewStreamWriter so all rows do not have to be kept in memory.
func Write(w io.Writer, wb *Workbook, f Format) error {
	if err := wb.Validate(); err != nil {
		return err
	}
	sw, err := newStreamWriter(w, f, wb, false)
	if err != nil {
		return err
	}
	for _, sh := range wb.Sheets {
		if err := sw.NextSheet(); err != nil {
			return err
		}
		for _, row := range sh.Rows {
			if err := sw.writeRow(row); err != nil {
				return err
			}
		}
	}
	return sw.Close()
}

// styleKey is a combination of a user style and the effective number format;
// date cells without a date format get one automatically, so one style can expand into several.
type styleKey struct {
	id  int
	fmt NumFmt
}

type styleTable struct {
	wb   *Workbook
	idx  map[styleKey]int // key -> 1-based index (0 is reserved for the default)
	keys []styleKey
}

func newStyleTable(wb *Workbook) *styleTable {
	return &styleTable{wb: wb, idx: map[styleKey]int{}}
}

func (t *styleTable) key(c Cell) styleKey {
	f := t.wb.style(c.Style).NumFmt
	if c.Kind == KindDate && !f.isDate() {
		if hasClock(c.Time) {
			f = NumFmtDateTime
		} else {
			f = NumFmtDate
		}
	}
	return styleKey{c.Style, f}
}

// get returns the style index for a cell; 0 means default.
func (t *styleTable) get(c Cell) int {
	k := t.key(c)
	if k.id == 0 && k.fmt == NumFmtGeneral {
		return 0
	}
	if i, ok := t.idx[k]; ok {
		return i
	}
	t.keys = append(t.keys, k)
	t.idx[k] = len(t.keys)
	return len(t.keys)
}

// ---- shared output helpers ----

type xw struct {
	*bufio.Writer
	err error
}

func newXW(w io.Writer) *xw { return &xw{Writer: bufio.NewWriterSize(w, 64<<10)} }

func (w *xw) s(s string) {
	if w.err == nil {
		_, w.err = w.WriteString(s)
	}
}

// esc writes escaped text; characters not allowed in XML (such as 0x00-0x08) are replaced with U+FFFD.
func (w *xw) esc(s string) {
	if w.err == nil {
		w.err = xml.EscapeText(w, []byte(s))
	}
}

func (w *xw) i(n int) { w.s(strconv.Itoa(n)) }

func (w *xw) f(v float64) { w.s(strconv.FormatFloat(v, 'g', -1, 64)) }

func (w *xw) done() error {
	if w.err != nil {
		return w.err
	}
	return w.Flush()
}

const xmlHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + "\n"

// ---- date conversion ----

var (
	excelEpoch1900 = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	excelEpoch1904 = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
)

// excelSerial converts a time (its wall-clock value) to an Excel serial number (1900 date system).
func excelSerial(t time.Time) float64 {
	u := wallUTC(t)
	secs := u.Unix() - excelEpoch1900.Unix()
	return float64(secs)/86400 + float64(u.Nanosecond())/86400e9
}

// fromExcelSerial converts an Excel serial number back to a time; it returns false when out of range.
func fromExcelSerial(v float64, date1904 bool, loc *time.Location) (time.Time, bool) {
	if v < 0 || v > 2958466 || v != v {
		return time.Time{}, false
	}
	epoch := excelEpoch1900
	if date1904 {
		epoch = excelEpoch1904
	} else if v < 61 {
		v++ // before 1900-03-01, undo the offset of the Lotus 1-2-3 leap-year bug
	}
	ms := int64(v*86400000 + 0.5)
	u := time.UnixMilli(epoch.UnixMilli() + ms).UTC()
	return time.Date(u.Year(), u.Month(), u.Day(), u.Hour(), u.Minute(), u.Second(), u.Nanosecond(), loc), true
}
