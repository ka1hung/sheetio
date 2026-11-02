package sheetio

import (
	"archive/zip"
	"fmt"
	"io"
	"strconv"
	"strings"
)

const (
	nsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	nsRel  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	nsPkg  = "http://schemas.openxmlformats.org/package/2006/relationships"
)

type xlsxSink struct {
	zw   *zip.Writer
	wb   *Workbook
	st   *styleTable
	x    *xw
	rows []int // rows written per sheet (needed for the auto filter range)
}

func (s *xlsxSink) put(name string, fn func(w *xw)) error {
	fw, err := s.zw.Create(name)
	if err != nil {
		return err
	}
	w := newXW(fw)
	fn(w)
	if err := w.done(); err != nil {
		return fmt.Errorf("sheetio: write %s: %w", name, err)
	}
	return nil
}

func newXLSXSink(out io.Writer, wb *Workbook, st *styleTable) (*xlsxSink, error) {
	s := &xlsxSink{zw: zip.NewWriter(out), wb: wb, st: st, rows: make([]int, len(wb.Sheets))}
	n := len(wb.Sheets)
	steps := []struct {
		name string
		fn   func(w *xw)
	}{
		{"[Content_Types].xml", func(w *xw) {
			w.s(xmlHeader)
			w.s(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`)
			w.s(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
			w.s(`<Default Extension="xml" ContentType="application/xml"/>`)
			w.s(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`)
			for i := 1; i <= n; i++ {
				w.s(`<Override PartName="/xl/worksheets/sheet`)
				w.i(i)
				w.s(`.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
			}
			w.s(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`)
			w.s(`</Types>`)
		}},
		{"_rels/.rels", func(w *xw) {
			w.s(xmlHeader)
			w.s(`<Relationships xmlns="` + nsPkg + `">`)
			w.s(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>`)
			w.s(`</Relationships>`)
		}},
		{"xl/_rels/workbook.xml.rels", func(w *xw) {
			w.s(xmlHeader)
			w.s(`<Relationships xmlns="` + nsPkg + `">`)
			for i := 1; i <= n; i++ {
				w.s(`<Relationship Id="rId`)
				w.i(i)
				w.s(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet`)
				w.i(i)
				w.s(`.xml"/>`)
			}
			w.s(`<Relationship Id="rId`)
			w.i(n + 1)
			w.s(`" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`)
			w.s(`</Relationships>`)
		}},
	}
	for _, st := range steps {
		if err := s.put(st.name, st.fn); err != nil {
			return nil, err
		}
	}
	return s, nil
}

func (s *xlsxSink) begin(i int) error {
	fw, err := s.zw.Create("xl/worksheets/sheet" + strconv.Itoa(i+1) + ".xml")
	if err != nil {
		return err
	}
	s.x = newXW(fw)
	writeXLSXSheetHead(s.x, s.wb.Sheets[i], i == 0)
	return s.x.err
}

func (s *xlsxSink) row(i, r int, cells []Cell) error {
	writeXLSXRow(s.x, r, cells, s.st, s.wb.Sheets[i].RowHeights[r])
	return s.x.err
}

func (s *xlsxSink) end(i, rows int) error {
	s.rows[i] = rows
	writeXLSXSheetTail(s.x, s.wb.Sheets[i], rows)
	return s.x.done()
}

func (s *xlsxSink) close() error {
	// workbook.xml is written last, once we know how far each auto filter extends
	if err := s.put("xl/workbook.xml", func(w *xw) { writeXLSXWorkbook(w, s.wb, s.rows) }); err != nil {
		return err
	}
	// styles.xml is written last, once we know which style combinations are used
	if err := s.put("xl/styles.xml", func(w *xw) { writeXLSXStyles(w, s.st) }); err != nil {
		return err
	}
	return s.zw.Close()
}

func writeXLSXWorkbook(w *xw, wb *Workbook, rows []int) {
	w.s(xmlHeader)
	w.s(`<workbook xmlns="` + nsMain + `" xmlns:r="` + nsRel + `">`)
	w.s(`<bookViews><workbookView activeTab="0"/></bookViews><sheets>`)
	for i, sh := range wb.Sheets {
		w.s(`<sheet name="`)
		w.esc(sh.Name)
		w.s(`" sheetId="`)
		w.i(i + 1)
		w.s(`" r:id="rId`)
		w.i(i + 1)
		w.s(`"/>`)
	}
	w.s(`</sheets>`)
	// Excel records each sheet's filter range in a hidden _FilterDatabase name
	started := false
	for i, sh := range wb.Sheets {
		f, ok := sh.filterRange(rows[i])
		if !ok {
			continue
		}
		if !started {
			w.s(`<definedNames>`)
			started = true
		}
		w.s(`<definedName name="_xlnm._FilterDatabase" localSheetId="`)
		w.i(i)
		w.s(`" hidden="1">`)
		w.esc(quoteSheetName(sh.Name) + "!" + absRef(f.Row1, f.Col1) + ":" + absRef(f.Row2, f.Col2))
		w.s(`</definedName>`)
	}
	if started {
		w.s(`</definedNames>`)
	}
	w.s(`</workbook>`)
}

// quoteSheetName quotes a sheet name for formulas and range references: always wrapped in single quotes, with embedded single quotes doubled.
func quoteSheetName(name string) string {
	return "'" + strings.ReplaceAll(name, "'", "''") + "'"
}

// absRef converts a 0-based row and column to $A$1.
func absRef(row, col int) string {
	return "$" + ColName(col) + "$" + strconv.Itoa(row+1)
}

func xlsxNumFmtID(f NumFmt) int {
	switch f {
	case NumFmtInteger:
		return 1
	case NumFmtDecimal2:
		return 2
	case NumFmtThousands:
		return 3
	case NumFmtDate:
		return 164
	case NumFmtDateTime:
		return 165
	}
	return 0
}

func writeXLSXStyles(w *xw, st *styleTable) {
	// collect the fill colors in use
	fillIdx := map[string]int{}
	var fills []string
	for _, k := range st.keys {
		c := st.wb.style(k.id).FillColor
		if c != "" {
			if _, ok := fillIdx[c]; !ok {
				fills = append(fills, c)
				fillIdx[c] = len(fills) + 1 // 0 and 1 are reserved by the spec
			}
		}
	}

	w.s(xmlHeader)
	w.s(`<styleSheet xmlns="` + nsMain + `">`)
	w.s(`<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy\-mm\-dd"/><numFmt numFmtId="165" formatCode="yyyy\-mm\-dd\ hh:mm:ss"/></numFmts>`)
	w.s(`<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>`)
	w.s(`<fills count="`)
	w.i(len(fills) + 2)
	w.s(`"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>`)
	for _, c := range fills {
		w.s(`<fill><patternFill patternType="solid"><fgColor rgb="FF`)
		w.s(c)
		w.s(`"/><bgColor indexed="64"/></patternFill></fill>`)
	}
	w.s(`</fills>`)
	w.s(`<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>`)
	w.s(`<border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders>`)
	w.s(`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`)
	w.s(`<cellXfs count="`)
	w.i(len(st.keys) + 1)
	w.s(`"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`)
	for _, k := range st.keys {
		s := st.wb.style(k.id)
		font, fill, border := 0, 0, 0
		if s.Bold {
			font = 1
		}
		if s.FillColor != "" {
			fill = fillIdx[s.FillColor]
		}
		if s.Border {
			border = 1
		}
		w.s(`<xf numFmtId="`)
		w.i(xlsxNumFmtID(k.fmt))
		w.s(`" fontId="`)
		w.i(font)
		w.s(`" fillId="`)
		w.i(fill)
		w.s(`" borderId="`)
		w.i(border)
		w.s(`" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"`)
		if s.Align != AlignDefault || s.Wrap {
			w.s(` applyAlignment="1"><alignment`)
			switch s.Align {
			case AlignLeft:
				w.s(` horizontal="left"`)
			case AlignCenter:
				w.s(` horizontal="center"`)
			case AlignRight:
				w.s(` horizontal="right"`)
			}
			if s.Wrap {
				w.s(` wrapText="1"`)
			}
			w.s(`/></xf>`)
		} else {
			w.s(`/>`)
		}
	}
	w.s(`</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`)
	w.s(`</styleSheet>`)
}

func writeXLSXSheetHead(w *xw, sh *Sheet, first bool) {
	w.s(xmlHeader)
	w.s(`<worksheet xmlns="` + nsMain + `" xmlns:r="` + nsRel + `">`)
	w.s(`<sheetViews><sheetView`)
	if first {
		w.s(` tabSelected="1"`)
	}
	w.s(` workbookViewId="0"`)
	if sh.FreezeRows > 0 || sh.FreezeCols > 0 {
		pane := "bottomRight"
		switch {
		case sh.FreezeCols == 0:
			pane = "bottomLeft"
		case sh.FreezeRows == 0:
			pane = "topRight"
		}
		w.s(`><pane`)
		if sh.FreezeCols > 0 {
			w.s(` xSplit="`)
			w.i(sh.FreezeCols)
			w.s(`"`)
		}
		if sh.FreezeRows > 0 {
			w.s(` ySplit="`)
			w.i(sh.FreezeRows)
			w.s(`"`)
		}
		w.s(` topLeftCell="` + CellRef(sh.FreezeRows, sh.FreezeCols) + `" activePane="` + pane + `" state="frozen"/>`)
		w.s(`<selection pane="` + pane + `"/></sheetView>`)
	} else {
		w.s(`/>`)
	}
	w.s(`</sheetViews>`)
	w.s(`<sheetFormatPr defaultRowHeight="15"/>`)
	hasCols := false
	for c, width := range sh.ColWidths {
		if width <= 0 {
			continue
		}
		if !hasCols {
			w.s(`<cols>`)
			hasCols = true
		}
		w.s(`<col min="`)
		w.i(c + 1)
		w.s(`" max="`)
		w.i(c + 1)
		w.s(`" width="`)
		w.f(width)
		w.s(`" customWidth="1"/>`)
	}
	if hasCols {
		w.s(`</cols>`)
	}
	w.s(`<sheetData>`)
}

func writeXLSXRow(w *xw, r int, row []Cell, st *styleTable, height float64) {
	started := false
	start := func() {
		w.s(`<row r="`)
		w.i(r + 1)
		if height > 0 {
			w.s(`" ht="`)
			w.f(height)
			w.s(`" customHeight="1`)
		}
		w.s(`">`)
		started = true
	}
	if height > 0 { // empty rows with a custom height are still written
		start()
	}
	for c, cell := range row {
		if cell.IsEmpty() && cell.Style == 0 {
			continue
		}
		if !started {
			start()
		}
		w.s(`<c r="`)
		w.s(CellRef(r, c))
		w.s(`"`)
		if s := st.get(cell); s != 0 {
			w.s(` s="`)
			w.i(s)
			w.s(`"`)
		}
		switch cell.Kind {
		case KindEmpty:
			w.s(`/>`)
		case KindString:
			// inlineStr is a string type, so text starting with = is never treated as a formula.
			w.s(` t="inlineStr"><is><t xml:space="preserve">`)
			w.esc(cell.Str)
			w.s(`</t></is></c>`)
		case KindNumber:
			w.s(`><v>`)
			w.f(cell.Num)
			w.s(`</v></c>`)
		case KindBool:
			if cell.Bool {
				w.s(` t="b"><v>1</v></c>`)
			} else {
				w.s(` t="b"><v>0</v></c>`)
			}
		case KindDate:
			w.s(`><v>`)
			w.f(excelSerial(cell.Time))
			w.s(`</v></c>`)
		}
	}
	if started {
		w.s(`</row>`)
	}
}

func writeXLSXSheetTail(w *xw, sh *Sheet, rows int) {
	w.s(`</sheetData>`)
	if f, ok := sh.filterRange(rows); ok {
		w.s(`<autoFilter ref="`)
		w.s(f.ref())
		w.s(`"/>`)
	}
	if len(sh.Merges) > 0 {
		w.s(`<mergeCells count="`)
		w.i(len(sh.Merges))
		w.s(`">`)
		for _, m := range sh.Merges {
			w.s(`<mergeCell ref="`)
			w.s(m.ref())
			w.s(`"/>`)
		}
		w.s(`</mergeCells>`)
	}
	w.s(`<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`)
	w.s(`</worksheet>`)
}
