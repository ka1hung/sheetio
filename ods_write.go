package sheetio

import (
	"archive/zip"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
	"time"
)

const odsNS = ` xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
	` xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"` +
	` xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"` +
	` xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"` +
	` xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"` +
	` xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"` +
	` xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"` +
	` office:version="1.2"`

// odsSink writes ODS as a stream.
//
// content.xml is laid out as "automatic styles, then each table (column definitions, then
// rows)", but the column count and the styles in use are only known after the rows. So each
// part is compressed as a separate deflate segment, and the segments are joined in the right
// order at the end:
//
//	[head: styles] [table 1 columns] [table 1 rows] [table 2 columns] [table 2 rows] ... [tail]
//
// Only the compressed segments are kept in memory.
type odsSink struct {
	z        *cleanZip
	wb       *Workbook
	st       *styleTable
	colStyle map[float64]string
	colOrder []float64
	rowStyle map[float64]string // row height -> row style name (collected while rows are written)
	rowOrder []float64
	filters  []string // auto filter ranges (ODF range addresses), written as table:database-range
	segs     []*segment

	// current sheet
	seg    *segment
	x      *xw
	maxCol int
	blank  int
	merges []Range // sorted by Row1
	next   int     // next merge in merges not yet active
	active []Range // merges covering the current row
	mRows  int     // number of rows the merges extend to
}

func newODSSink(out io.Writer, wb *Workbook, st *styleTable) (*odsSink, error) {
	s := &odsSink{z: newCleanZip(out, time.Now()), wb: wb, st: st, colStyle: map[float64]string{0: "co0"}, rowStyle: map[float64]string{}}
	for _, sh := range wb.Sheets {
		for _, w := range sh.ColWidths {
			if _, ok := s.colStyle[w]; !ok && w > 0 {
				s.colStyle[w] = "co" + strconv.Itoa(len(s.colStyle))
				s.colOrder = append(s.colOrder, w)
			}
		}
	}
	// The spec requires mimetype to be the first entry, stored uncompressed, with no extra field and no data descriptor.
	if err := s.z.add("mimetype", zip.Store, func(w io.Writer) error {
		_, err := io.WriteString(w, mimeODS)
		return err
	}); err != nil {
		return nil, err
	}
	put := func(name string, fn func(w *xw)) error {
		return s.z.add(name, zip.Deflate, func(fw io.Writer) error {
			w := newXW(fw)
			fn(w)
			if err := w.done(); err != nil {
				return fmt.Errorf("sheetio: write %s: %w", name, err)
			}
			return nil
		})
	}
	if err := put("META-INF/manifest.xml", func(w *xw) {
		w.s(xmlHeader)
		w.s(`<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">`)
		w.s(`<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="` + mimeODS + `"/>`)
		w.s(`<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>`)
		w.s(`<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>`)
		w.s(`<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>`)
		w.s(`<manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>`)
		w.s(`</manifest:manifest>`)
	}); err != nil {
		return nil, err
	}
	if err := put("meta.xml", func(w *xw) {
		w.s(xmlHeader)
		w.s(`<office:document-meta` + odsNS + `><office:meta><meta:generator>sheetio</meta:generator></office:meta></office:document-meta>`)
	}); err != nil {
		return nil, err
	}
	if err := put("styles.xml", func(w *xw) {
		w.s(xmlHeader)
		w.s(`<office:document-styles` + odsNS + `><office:styles>`)
		w.s(`<style:style style:name="Default" style:family="table-cell"/>`)
		w.s(`</office:styles></office:document-styles>`)
	}); err != nil {
		return nil, err
	}
	return s, nil
}

func (s *odsSink) begin(i int) error {
	sh := s.wb.Sheets[i]
	s.seg = newSegment()
	s.x = newXW(s.seg)
	s.blank, s.next, s.active, s.mRows = 0, 0, s.active[:0], 0
	s.maxCol = len(sh.ColWidths)
	s.merges = append(s.merges[:0], sh.Merges...)
	sort.Slice(s.merges, func(a, b int) bool { return s.merges[a].Row1 < s.merges[b].Row1 })
	for _, m := range s.merges {
		if m.Col2+1 > s.maxCol {
			s.maxCol = m.Col2 + 1
		}
		if m.Row2+1 > s.mRows {
			s.mRows = m.Row2 + 1
		}
	}
	return nil
}

func (s *odsSink) flushBlank() {
	if s.blank == 0 {
		return
	}
	w := s.x
	w.s(`<table:table-row`)
	if s.blank > 1 {
		w.s(` table:number-rows-repeated="`)
		w.i(s.blank)
		w.s(`"`)
	}
	w.s(`><table:table-cell/></table:table-row>`)
	s.blank = 0
}

// rowStyleName returns the row style name for the row height, or "" when no height is set.
func (s *odsSink) rowStyleName(i, r int) string {
	h := s.wb.Sheets[i].RowHeights[r]
	if h <= 0 {
		return ""
	}
	name, ok := s.rowStyle[h]
	if !ok {
		name = "ro" + strconv.Itoa(len(s.rowStyle)+1)
		s.rowStyle[h] = name
		s.rowOrder = append(s.rowOrder, h)
	}
	return name
}

func (s *odsSink) row(i, r int, cells []Cell) error {
	// update the merges covering the current row (rows are written in order)
	keep := s.active[:0]
	for _, m := range s.active {
		if m.Row2 >= r {
			keep = append(keep, m)
		}
	}
	s.active = keep
	for s.next < len(s.merges) && s.merges[s.next].Row1 <= r {
		if s.merges[s.next].Row2 >= r {
			s.active = append(s.active, s.merges[s.next])
		}
		s.next++
	}
	sort.Slice(s.active, func(a, b int) bool { return s.active[a].Col1 < s.active[b].Col1 })

	width := len(cells)
	for len(cells) > 0 && width > 0 && cells[width-1].IsEmpty() && cells[width-1].Style == 0 {
		width--
	}
	if len(s.active) > 0 && s.active[len(s.active)-1].Col2+1 > width {
		width = s.active[len(s.active)-1].Col2 + 1
	}
	rs := s.rowStyleName(i, r)
	if width == 0 && rs == "" {
		s.blank++
		return nil
	}
	if width > s.maxCol {
		s.maxCol = width
	}
	s.flushBlank()

	w := s.x
	w.s(`<table:table-row`)
	if rs != "" {
		w.s(` table:style-name="` + rs + `"`)
	}
	w.s(`>`)
	if width == 0 { // empty row with a custom height
		w.s(`<table:table-cell/></table:table-row>`)
		return w.err
	}
	emptyRun := 0
	flushEmpty := func() {
		if emptyRun == 0 {
			return
		}
		w.s(`<table:table-cell`)
		if emptyRun > 1 {
			w.s(` table:number-columns-repeated="`)
			w.i(emptyRun)
			w.s(`"`)
		}
		w.s(`/>`)
		emptyRun = 0
	}
	ai := 0
	for c := 0; c < width; {
		if ai < len(s.active) && s.active[ai].Col1 == c {
			m := s.active[ai]
			ai++
			flushEmpty()
			covered := m.Col2 - m.Col1 + 1
			if m.Row1 == r {
				var cell Cell
				if c < len(cells) {
					cell = cells[c]
				}
				writeODSCell(w, cell, s.st, m, true)
				covered--
			}
			if covered > 0 {
				w.s(`<table:covered-table-cell`)
				if covered > 1 {
					w.s(` table:number-columns-repeated="`)
					w.i(covered)
					w.s(`"`)
				}
				w.s(`/>`)
			}
			c = m.Col2 + 1
			continue
		}
		var cell Cell
		if c < len(cells) {
			cell = cells[c]
		}
		if cell.IsEmpty() && cell.Style == 0 {
			emptyRun++
		} else {
			flushEmpty()
			writeODSCell(w, cell, s.st, Range{}, false)
		}
		c++
	}
	flushEmpty()
	w.s(`</table:table-row>`)
	return w.err
}

func (s *odsSink) end(i, rows int) error {
	// merges that extend past the data rows must still be written
	for r := rows; r < s.mRows; r++ {
		if err := s.row(i, r, nil); err != nil {
			return err
		}
	}
	if s.maxCol == 0 { // completely empty sheet
		s.x.s(`<table:table-row><table:table-cell/></table:table-row>`)
		s.maxCol = 1
	}
	// trailing empty rows are not written
	s.x.s(`</table:table>`)
	if f, ok := s.wb.Sheets[i].filterRange(rows); ok {
		q := quoteSheetName(s.wb.Sheets[i].Name)
		s.filters = append(s.filters, q+"."+CellRef(f.Row1, f.Col1)+":"+q+"."+CellRef(f.Row2, f.Col2))
	}
	if err := s.x.done(); err != nil {
		return err
	}
	if err := s.seg.finish(false); err != nil {
		return err
	}

	// add the table header and column definitions (the column count is known now)
	sh := s.wb.Sheets[i]
	hs := newSegment()
	w := newXW(hs)
	w.s(`<table:table table:name="`)
	w.esc(sh.Name)
	w.s(`">`)
	widthAt := func(c int) float64 {
		if c < len(sh.ColWidths) {
			return sh.ColWidths[c]
		}
		return 0
	}
	for c := 0; c < s.maxCol; {
		j := c + 1
		for j < s.maxCol && widthAt(j) == widthAt(c) {
			j++
		}
		w.s(`<table:table-column table:style-name="`)
		w.s(s.colStyle[widthAt(c)])
		w.s(`"`)
		if j-c > 1 {
			w.s(` table:number-columns-repeated="`)
			w.i(j - c)
			w.s(`"`)
		}
		w.s(` table:default-cell-style-name="Default"/>`)
		c = j
	}
	if err := w.done(); err != nil {
		return err
	}
	if err := hs.finish(false); err != nil {
		return err
	}
	s.segs = append(s.segs, hs, s.seg)
	s.seg, s.x = nil, nil
	return nil
}

func (s *odsSink) close() error {
	head := newSegment()
	w := newXW(head)
	w.s(xmlHeader)
	w.s(`<office:document-content` + odsNS + `>`)
	w.s(`<office:automatic-styles>`)
	// number and date formats
	w.s(`<number:number-style style:name="N1"><number:number number:decimal-places="0" number:min-integer-digits="1"/></number:number-style>`)
	w.s(`<number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:number-style>`)
	w.s(`<number:number-style style:name="N3"><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true"/></number:number-style>`)
	dateParts := `<number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/>`
	w.s(`<number:date-style style:name="N4">` + dateParts + `</number:date-style>`)
	w.s(`<number:date-style style:name="N5">` + dateParts +
		`<number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:date-style>`)

	w.s(`<style:style style:name="co0" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="2.258cm"/></style:style>`)
	for _, width := range s.colOrder {
		w.s(`<style:style style:name="` + s.colStyle[width] + `" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="` + odsColWidth(width) + `"/></style:style>`)
	}
	for _, h := range s.rowOrder {
		w.s(`<style:style style:name="` + s.rowStyle[h] + `" style:family="table-row"><style:table-row-properties style:row-height="` +
			strconv.FormatFloat(h*2.54/72, 'f', 3, 64) + `cm" style:use-optimal-row-height="false" fo:break-before="auto"/></style:style>`)
	}
	wb, st := s.wb, s.st
	// cell styles
	for i, k := range st.keys {
		s := wb.style(k.id)
		w.s(`<style:style style:name="ce`)
		w.i(i + 1)
		w.s(`" style:family="table-cell" style:parent-style-name="Default"`)
		if ds := odsDataStyle(k.fmt); ds != "" {
			w.s(` style:data-style-name="` + ds + `"`)
		}
		w.s(`>`)
		if s.FillColor != "" || s.Border || s.Wrap || s.Align != AlignDefault {
			w.s(`<style:table-cell-properties`)
			if s.FillColor != "" {
				w.s(` fo:background-color="#` + strings.ToLower(s.FillColor) + `"`)
			}
			if s.Border {
				w.s(` fo:border="0.06pt solid #000000"`)
			}
			if s.Wrap {
				w.s(` fo:wrap-option="wrap"`)
			}
			if s.Align != AlignDefault {
				w.s(` style:text-align-source="fix"`)
			}
			w.s(`/>`)
		}
		switch s.Align {
		case AlignLeft:
			w.s(`<style:paragraph-properties fo:text-align="start"/>`)
		case AlignCenter:
			w.s(`<style:paragraph-properties fo:text-align="center"/>`)
		case AlignRight:
			w.s(`<style:paragraph-properties fo:text-align="end"/>`)
		}
		if s.Bold {
			w.s(`<style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"/>`)
		}
		w.s(`</style:style>`)
	}
	w.s(`</office:automatic-styles><office:body><office:spreadsheet>`)
	if err := w.done(); err != nil {
		return err
	}
	if err := head.finish(false); err != nil {
		return err
	}
	tail := newSegment()
	tw := newXW(tail)
	if len(s.filters) > 0 {
		tw.s(`<table:database-ranges>`)
		for _, f := range s.filters {
			// no table:name: Excel ignores the filter when the range has a name (Excel's own ods output omits it too)
			tw.s(`<table:database-range table:target-range-address="`)
			tw.esc(f)
			tw.s(`" table:display-filter-buttons="true"/>`)
		}
		tw.s(`</table:database-ranges>`)
	}
	tw.s(`</office:spreadsheet></office:body></office:document-content>`)
	if err := tw.done(); err != nil {
		return err
	}
	if err := tail.finish(true); err != nil {
		return err
	}
	segs := append([]*segment{head}, s.segs...)
	segs = append(segs, tail)
	s.segs = nil
	if err := s.z.addSegments("content.xml", segs); err != nil {
		return err
	}
	if err := s.z.add("settings.xml", zip.Deflate, func(fw io.Writer) error {
		w := newXW(fw)
		writeODSSettings(w, s.wb)
		return w.done()
	}); err != nil {
		return err
	}
	return s.z.close()
}

// writeODSSettings writes freeze panes to the view settings in settings.xml (Excel ignores them when opening ods).
func writeODSSettings(w *xw, wb *Workbook) {
	w.s(xmlHeader)
	w.s(`<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"` +
		` xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.2"><office:settings>`)
	frozen := false
	for _, sh := range wb.Sheets {
		if sh.FreezeRows > 0 || sh.FreezeCols > 0 {
			frozen = true
		}
	}
	if frozen {
		item := func(name, typ string, v int) {
			w.s(`<config:config-item config:name="` + name + `" config:type="` + typ + `">`)
			w.i(v)
			w.s(`</config:config-item>`)
		}
		w.s(`<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>`)
		w.s(`<config:config-item config:name="ViewId" config:type="string">view1</config:config-item>`)
		w.s(`<config:config-item-map-named config:name="Tables">`)
		for _, sh := range wb.Sheets {
			if sh.FreezeRows == 0 && sh.FreezeCols == 0 {
				continue
			}
			mode := func(n int) int {
				if n > 0 {
					return 2 // 2 = frozen
				}
				return 0
			}
			active := 2 // 2 = bottom left (rows only); 3 = bottom right
			if sh.FreezeCols > 0 {
				active = 3
				if sh.FreezeRows == 0 {
					active = 1 // 1 = top right (columns only)
				}
			}
			w.s(`<config:config-item-map-entry config:name="`)
			w.esc(sh.Name)
			w.s(`">`)
			item("HorizontalSplitMode", "short", mode(sh.FreezeCols))
			item("VerticalSplitMode", "short", mode(sh.FreezeRows))
			item("HorizontalSplitPosition", "int", sh.FreezeCols)
			item("VerticalSplitPosition", "int", sh.FreezeRows)
			item("ActiveSplitRange", "short", active)
			item("PositionLeft", "int", 0)
			item("PositionRight", "int", sh.FreezeCols)
			item("PositionTop", "int", 0)
			item("PositionBottom", "int", sh.FreezeRows)
			w.s(`</config:config-item-map-entry>`)
		}
		w.s(`</config:config-item-map-named></config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set>`)
	}
	w.s(`</office:settings></office:document-settings>`)
}

func odsDataStyle(f NumFmt) string {
	switch f {
	case NumFmtInteger:
		return "N1"
	case NumFmtDecimal2:
		return "N2"
	case NumFmtThousands:
		return "N3"
	case NumFmtDate:
		return "N4"
	case NumFmtDateTime:
		return "N5"
	}
	return ""
}

// Column width: Excel "characters" converted roughly to centimetres (the default 8.43 characters is about 1.69 cm).
func odsColWidth(chars float64) string {
	cm := chars*0.19 + 0.09
	return strconv.FormatFloat(cm, 'f', 3, 64) + "cm"
}

func writeODSCell(w *xw, c Cell, st *styleTable, m Range, isSpan bool) {
	w.s(`<table:table-cell`)
	if s := st.get(c); s != 0 {
		w.s(` table:style-name="ce`)
		w.i(s)
		w.s(`"`)
	}
	if isSpan {
		w.s(` table:number-columns-spanned="`)
		w.i(m.Col2 - m.Col1 + 1)
		w.s(`" table:number-rows-spanned="`)
		w.i(m.Row2 - m.Row1 + 1)
		w.s(`"`)
	}
	switch c.Kind {
	case KindEmpty:
		w.s(`/>`)
		return
	case KindString:
		w.s(` office:value-type="string">`)
		writeODSText(w, c.Str)
	case KindNumber:
		// ODS uses plain decimal notation, not exponent forms like 1e+15
		v := strconv.FormatFloat(c.Num, 'f', -1, 64)
		w.s(` office:value-type="float" office:value="` + v + `"><text:p>` + v)
		w.s(`</text:p>`)
	case KindBool:
		v := "false"
		if c.Bool {
			v = "true"
		}
		w.s(` office:value-type="boolean" office:boolean-value="` + v + `"><text:p>` + strings.ToUpper(v) + `</text:p>`)
	case KindDate:
		u := wallUTC(c.Time)
		w.s(` office:value-type="date" office:date-value="`)
		if hasClock(u) {
			w.s(u.Format("2006-01-02T15:04:05"))
		} else {
			w.s(u.Format("2006-01-02"))
		}
		w.s(`"><text:p>`)
		w.s(c.String())
		w.s(`</text:p>`)
	}
	w.s(`</table:table-cell>`)
}

// writeODSText writes text following the ODF whitespace rules: line breaks become separate text:p elements,
// leading or repeated spaces become text:s, and tabs become text:tab.
func writeODSText(w *xw, s string) {
	s = strings.ReplaceAll(s, "\r\n", "\n")
	for _, line := range strings.Split(s, "\n") {
		w.s(`<text:p>`)
		var seg strings.Builder
		flush := func() {
			if seg.Len() > 0 {
				w.esc(seg.String())
				seg.Reset()
			}
		}
		rs := []rune(line)
		for i := 0; i < len(rs); {
			switch rs[i] {
			case ' ':
				j := i
				for j < len(rs) && rs[j] == ' ' {
					j++
				}
				n := j - i
				// only a single space between two ordinary characters can be written as is;
				// leading, trailing or repeated spaces always use text:s so XML whitespace handling does not drop them.
				if n == 1 && i > 0 && j < len(rs) && rs[i-1] != '\t' && rs[j] != '\t' {
					seg.WriteRune(' ')
				} else {
					flush()
					w.s(`<text:s`)
					if n > 1 {
						w.s(` text:c="`)
						w.i(n)
						w.s(`"`)
					}
					w.s(`/>`)
				}
				i = j
			case '\t':
				flush()
				w.s(`<text:tab/>`)
				i++
			case '\r':
				i++
			default:
				seg.WriteRune(rs[i])
				i++
			}
		}
		flush()
		w.s(`</text:p>`)
	}
}
