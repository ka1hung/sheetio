package sheetio

import (
	"archive/zip"
	"bytes"
	"errors"
	"io"
	"strings"
	"testing"
)

// layoutWorkbook has a header row and 5 data rows with an empty row that has a custom height; the sheet name contains ' to test escaping.
func layoutWorkbook() *Workbook {
	wb := New()
	s := wb.AddSheet("報表 'A' 月")
	s.AddRow(Str("編號"), Str("名稱"), Str("數量"))
	s.AddRow(Int(1), Str("a"), Num(1.5))
	s.AddRow(Int(2), Str("b"), Num(2.5))
	s.AddRow() // empty row
	s.AddRow(Int(4), Str("d"), Num(4.5))
	s.AddRow(Int(5), Str("e"), Num(5.5))
	s.Freeze(1, 1)
	s.SetAutoFilter(0, 0, 2)
	s.SetRowHeight(0, 30)
	s.SetRowHeight(3, 25)
	wb.AddSheet("Plain").AddRow(Str("x"))
	return wb
}

// zipText returns the content of a zip entry.
func zipText(t *testing.T, b []byte, name string) string {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				t.Fatal(err)
			}
			defer rc.Close()
			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatal(err)
			}
			return string(data)
		}
	}
	t.Fatalf("%s not found", name)
	return ""
}

func mustContain(t *testing.T, where, s string, subs ...string) {
	t.Helper()
	for _, sub := range subs {
		if !strings.Contains(s, sub) {
			t.Errorf("%s: missing %q", where, sub)
		}
	}
}

func writeBytes(t *testing.T, wb *Workbook, f Format) []byte {
	t.Helper()
	var buf bytes.Buffer
	if err := Write(&buf, wb, f); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func TestLayoutXLSX(t *testing.T) {
	b := writeBytes(t, layoutWorkbook(), FormatXLSX)
	sheet := zipText(t, b, "xl/worksheets/sheet1.xml")
	mustContain(t, "sheet1", sheet,
		`<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>`,
		`<row r="1" ht="30" customHeight="1">`,
		`<row r="4" ht="25" customHeight="1"></row>`,
		`</sheetData><autoFilter ref="A1:C6"/>`)
	mustContain(t, "workbook", zipText(t, b, "xl/workbook.xml"),
		`<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">&#39;報表 &#39;&#39;A&#39;&#39; 月&#39;!$A$1:$C$6</definedName>`)
	if s := zipText(t, b, "xl/worksheets/sheet2.xml"); strings.Contains(s, "<pane") || strings.Contains(s, "autoFilter") {
		t.Errorf("sheet2 should have no layout: %s", s)
	}

	// the values read back are unaffected
	wb, _, err := Read(bytes.NewReader(b), int64(len(b)), nil)
	if err != nil {
		t.Fatal(err)
	}
	if got := wb.Sheets[0].Strings(); len(got) != 6 || got[5][1] != "e" || len(got[3]) != 0 {
		t.Errorf("read back = %q", got)
	}
}

func TestLayoutODS(t *testing.T) {
	b := writeBytes(t, layoutWorkbook(), FormatODS)
	content := zipText(t, b, "content.xml")
	mustContain(t, "content", content,
		`<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="1.058cm"`,
		`<style:style style:name="ro2" style:family="table-row"><style:table-row-properties style:row-height="0.882cm"`,
		`<table:table-row table:style-name="ro1">`,
		`<table:table-row table:style-name="ro2"><table:table-cell/></table:table-row>`,
		`table:target-range-address="&#39;報表 &#39;&#39;A&#39;&#39; 月&#39;.A1:&#39;報表 &#39;&#39;A&#39;&#39; 月&#39;.C6" table:display-filter-buttons="true"`)
	mustContain(t, "settings", zipText(t, b, "settings.xml"),
		`config:name="報表 &#39;A&#39; 月"`,
		`<config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item>`,
		`<config:config-item config:name="VerticalSplitPosition" config:type="int">1</config:config-item>`)
	mustContain(t, "manifest", zipText(t, b, "META-INF/manifest.xml"), `manifest:full-path="settings.xml"`)

	wb, _, err := Read(bytes.NewReader(b), int64(len(b)), nil)
	if err != nil {
		t.Fatal(err)
	}
	if got := wb.Sheets[0].Strings(); len(got) != 6 || got[5][1] != "e" || len(got[3]) != 0 {
		t.Errorf("read back = %q", got)
	}
}

// Without layout settings the output must not contain any layout elements.
func TestLayoutNone(t *testing.T) {
	wb := New()
	wb.AddSheet("S").AddRow(Str("a"))
	x := writeBytes(t, wb, FormatXLSX)
	for _, s := range []string{zipText(t, x, "xl/worksheets/sheet1.xml"), zipText(t, x, "xl/workbook.xml")} {
		if strings.Contains(s, "pane") || strings.Contains(s, "autoFilter") || strings.Contains(s, "definedName") || strings.Contains(s, "customHeight") {
			t.Errorf("unexpected layout: %s", s)
		}
	}
	o := writeBytes(t, wb, FormatODS)
	if s := zipText(t, o, "content.xml"); strings.Contains(s, "database-range") || strings.Contains(s, "table-row-properties") {
		t.Errorf("unexpected layout: %s", s)
	}
	if s := zipText(t, o, "settings.xml"); strings.Contains(s, "config-item") {
		t.Errorf("unexpected settings: %s", s)
	}
}

// Streaming: a filter given on the header row only extends to the last written row.
func TestLayoutStream(t *testing.T) {
	for _, f := range []Format{FormatXLSX, FormatODS} {
		var buf bytes.Buffer
		sw, err := NewStreamWriter(&buf, f, nil, []SheetDef{{
			Name: "Log", FreezeRows: 1, AutoFilter: &Range{0, 0, 0, 1},
			RowHeights: map[int]float64{0: 20},
		}})
		if err != nil {
			t.Fatal(err)
		}
		for i := 0; i < 100; i++ {
			if err := sw.WriteRow(Int(int64(i)), Str("x")); err != nil {
				t.Fatal(err)
			}
		}
		if err := sw.Close(); err != nil {
			t.Fatal(err)
		}
		if f == FormatXLSX {
			sheet := zipText(t, buf.Bytes(), "xl/worksheets/sheet1.xml")
			mustContain(t, "stream xlsx", sheet, `<autoFilter ref="A1:B100"/>`, `ySplit="1" topLeftCell="A2" activePane="bottomLeft"`)
		} else {
			mustContain(t, "stream ods", zipText(t, buf.Bytes(), "content.xml"), `&#39;Log&#39;.A1:&#39;Log&#39;.B100`)
			mustContain(t, "stream ods", zipText(t, buf.Bytes(), "settings.xml"),
				`<config:config-item config:name="HorizontalSplitMode" config:type="short">0</config:config-item>`,
				`<config:config-item config:name="ActiveSplitRange" config:type="short">2</config:config-item>`)
		}
	}
}

func TestLayoutValidate(t *testing.T) {
	cases := map[string]func(s *Sheet){
		"row height 0":      func(s *Sheet) { s.SetRowHeight(0, 0) },
		"row height 410":    func(s *Sheet) { s.SetRowHeight(0, 410) },
		"row height row -1": func(s *Sheet) { s.SetRowHeight(-1, 20) },
		"freeze negative":   func(s *Sheet) { s.Freeze(-1, 0) },
		"freeze too many":   func(s *Sheet) { s.Freeze(0, MaxSheetCols) },
		"filter reversed":   func(s *Sheet) { s.SetAutoFilter(0, 3, 1) },
		"filter negative":   func(s *Sheet) { s.SetAutoFilter(-1, 0, 1) },
	}
	for name, fn := range cases {
		wb := New()
		s := wb.AddSheet("S")
		s.AddRow(Str("a"))
		fn(s)
		if err := wb.Validate(); !errors.Is(err, ErrInvalid) {
			t.Errorf("%s: want ErrInvalid, got %v", name, err)
		}
	}
}
