package sheetio

import (
	"archive/zip"
	"bytes"
	"errors"
	"math"
	"os"
	"strings"
	"testing"
	"time"
)

var tpe = time.FixedZone("Asia/Taipei", 8*3600)

func sampleWorkbook() *Workbook {
	wb := New()
	head := wb.AddStyle(Style{Bold: true, FillColor: "D9E1F2", Border: true, Align: AlignCenter})
	money := wb.AddStyle(Style{NumFmt: NumFmtThousands, Border: true})
	wrap := wb.AddStyle(Style{Wrap: true})

	s := wb.AddSheet("員工清單")
	s.SetColWidth(0, 12)
	s.SetColWidth(1, 20)
	s.SetColWidth(4, 22)
	s.AddRow(Str("工號").WithStyle(head), Str("姓名").WithStyle(head), Str("薪資").WithStyle(head),
		Str("在職").WithStyle(head), Str("到職時間").WithStyle(head), Str("備註").WithStyle(head))
	s.AddRow(Str("A001"), Str("王小明"), Int(52000).WithStyle(money), Bool(true),
		Date(time.Date(2020, 3, 15, 0, 0, 0, 0, tpe)), Str("第一行\n第二行").WithStyle(wrap))
	s.AddRow(Str("A002"), Str("  前導空白 與  兩個空白 "), Num(48000.5), Bool(false),
		Date(time.Date(2026, 11, 1, 23, 5, 9, 0, tpe)), Str("=1+1 不應被執行"))
	s.AddRow(Str("A003"), Str("控制字元\x01測試"), Num(-0.125), Empty(), Empty(), Str("tab\t後"))
	s.AddRow() // empty row
	s.AddRow(Str("最後一列"), Empty(), Num(1e15))

	s2 := wb.AddSheet("Summary")
	s2.Merge(0, 0, 0, 2)
	s2.AddRow(Str("合併標題").WithStyle(head))
	s2.AddRow(Str("a"), Str("b"), Str("c"))
	s2.AddRow(Num(1), Num(2), Num(3))

	wb.AddSheet("Empty")
	return wb
}

// Values expected after reading back (control characters become U+FFFD).
func expectedRows() [][]string {
	return [][]string{
		{"工號", "姓名", "薪資", "在職", "到職時間", "備註"},
		{"A001", "王小明", "52000", "TRUE", "2020-03-15", "第一行\n第二行"},
		{"A002", "  前導空白 與  兩個空白 ", "48000.5", "FALSE", "2026-11-01 23:05:09", "=1+1 不應被執行"},
		{"A003", "控制字元�測試", "-0.125", "", "", "tab\t後"},
		nil,
		{"最後一列", "", "1000000000000000"},
	}
}

func roundTrip(t *testing.T, f Format) *Workbook {
	t.Helper()
	var buf bytes.Buffer
	if err := Write(&buf, sampleWorkbook(), f); err != nil {
		t.Fatalf("write %v: %v", f, err)
	}
	if os.Getenv("SHEETIO_DUMP") != "" {
		_ = os.WriteFile("testdata_out/sample"+f.Ext(), buf.Bytes(), 0o644)
	}
	wb, got, err := Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()), &ReadOptions{Location: tpe})
	if err != nil {
		t.Fatalf("read %v: %v", f, err)
	}
	if got != f {
		t.Fatalf("detected %v, want %v", got, f)
	}
	return wb
}

func TestRoundTrip(t *testing.T) {
	for _, f := range []Format{FormatXLSX, FormatODS} {
		t.Run(f.String(), func(t *testing.T) {
			wb := roundTrip(t, f)
			if len(wb.Sheets) != 3 {
				t.Fatalf("sheets = %d", len(wb.Sheets))
			}
			if wb.Sheets[0].Name != "員工清單" || wb.Sheets[1].Name != "Summary" || wb.Sheets[2].Name != "Empty" {
				t.Fatalf("names: %q %q %q", wb.Sheets[0].Name, wb.Sheets[1].Name, wb.Sheets[2].Name)
			}
			got := wb.Sheets[0].Strings()
			want := expectedRows()
			if len(got) != len(want) {
				t.Fatalf("rows = %d, want %d: %q", len(got), len(want), got)
			}
			for r := range want {
				if len(got[r]) != len(want[r]) {
					t.Errorf("row %d = %q, want %q", r, got[r], want[r])
					continue
				}
				for c := range want[r] {
					if got[r][c] != want[r][c] {
						t.Errorf("%s = %q, want %q", CellRef(r, c), got[r][c], want[r][c])
					}
				}
			}
			// type checks
			sh := wb.Sheets[0]
			if k := sh.Cell(1, 2).Kind; k != KindNumber {
				t.Errorf("C2 kind = %v", k)
			}
			if k := sh.Cell(1, 3).Kind; k != KindBool {
				t.Errorf("D2 kind = %v", k)
			}
			if c := sh.Cell(2, 4); c.Kind != KindDate || !c.Time.Equal(time.Date(2026, 11, 1, 23, 5, 9, 0, tpe)) {
				t.Errorf("E3 = %+v", c)
			}
			s2 := wb.Sheets[1].Strings()
			if s2[0][0] != "合併標題" || s2[2][2] != "3" {
				t.Errorf("summary = %q", s2)
			}
			if len(wb.Sheets[2].Rows) != 0 {
				t.Errorf("empty sheet has rows: %q", wb.Sheets[2].Strings())
			}
		})
	}
}

func TestExcelSerial(t *testing.T) {
	cases := []struct {
		t time.Time
		v float64
	}{
		{time.Date(1900, 3, 1, 0, 0, 0, 0, time.UTC), 61},
		{time.Date(2000, 1, 1, 0, 0, 0, 0, time.UTC), 36526},
		{time.Date(2000, 1, 1, 12, 0, 0, 0, tpe), 36526.5}, // wall-clock time, unaffected by the time zone
		{time.Date(9999, 12, 31, 0, 0, 0, 0, time.UTC), 2958465},
	}
	for _, c := range cases {
		if got := excelSerial(c.t); math.Abs(got-c.v) > 1e-9 {
			t.Errorf("serial(%v) = %v, want %v", c.t, got, c.v)
		}
		back, ok := fromExcelSerial(c.v, false, c.t.Location())
		if !ok || !back.Equal(c.t) {
			t.Errorf("fromSerial(%v) = %v, want %v", c.v, back, c.t)
		}
	}
	if d, _ := fromExcelSerial(1, false, time.UTC); d.Format("2006-01-02") != "1900-01-01" {
		t.Errorf("serial 1 = %v", d)
	}
	if d, _ := fromExcelSerial(0, true, time.UTC); d.Format("2006-01-02") != "1904-01-01" {
		t.Errorf("1904 serial 0 = %v", d)
	}
}

func TestDateFormatDetection(t *testing.T) {
	yes := []string{"yyyy-mm-dd", "m/d/yy", "[$-404]e/m/d", `yyyy"年"m"月"d"日"`, "hh:mm:ss", "[h]:mm", "dd-mmm"}
	no := []string{"General", "0.00", "#,##0", `0.00E+00`, `"days"0`, "[Red]0.00", `#,##0" 元"`, "@"}
	for _, s := range yes {
		if !isDateFormatCode(s) {
			t.Errorf("%q should be date", s)
		}
	}
	for _, s := range no {
		if isDateFormatCode(s) {
			t.Errorf("%q should not be date", s)
		}
	}
}

func TestValidate(t *testing.T) {
	bad := []func(wb *Workbook){
		func(wb *Workbook) { wb.AddSheet("a/b") },
		func(wb *Workbook) { wb.AddSheet(strings.Repeat("長", 32)) },
		func(wb *Workbook) { wb.AddSheet("Dup"); wb.AddSheet("dup") },
		func(wb *Workbook) { wb.AddSheet("x").AddRow(Num(math.NaN())) },
		func(wb *Workbook) { wb.AddSheet("x").AddRow(Str(strings.Repeat("a", MaxCellChars+1))) },
		func(wb *Workbook) { wb.AddSheet("x").AddRow(Str("a").WithStyle(9)) },
		func(wb *Workbook) { wb.AddSheet("x").AddRow(Date(time.Date(1899, 1, 1, 0, 0, 0, 0, time.UTC))) },
		func(wb *Workbook) { s := wb.AddSheet("x"); s.Merge(0, 0, 1, 1); s.Merge(1, 1, 2, 2) },
		func(wb *Workbook) { wb.AddStyle(Style{FillColor: "red"}); wb.AddSheet("x") },
	}
	for i, fn := range bad {
		wb := New()
		fn(wb)
		if err := Write(&bytes.Buffer{}, wb, FormatXLSX); !errors.Is(err, ErrInvalid) {
			t.Errorf("case %d: err = %v, want ErrInvalid", i, err)
		}
	}
	if err := Write(&bytes.Buffer{}, New(), FormatODS); !errors.Is(err, ErrInvalid) {
		t.Errorf("empty workbook: %v", err)
	}
}

// ---- malicious / damaged files ----

type entry struct{ name, body string }

func makeZip(t testing.TB, entries ...entry) []byte {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for _, e := range entries {
		w, err := zw.Create(e.name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(e.body)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func xlsxWith(t testing.TB, sheetXML, sst string) []byte {
	es := []entry{
		{"[Content_Types].xml", `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`},
		{"_rels/.rels", `<Relationships xmlns="` + nsPkg + `"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`},
		{"xl/workbook.xml", `<workbook xmlns="` + nsMain + `" xmlns:r="` + nsRel + `"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`},
		{"xl/_rels/workbook.xml.rels", `<Relationships xmlns="` + nsPkg + `"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="/xl/sharedStrings.xml"/></Relationships>`},
		{"xl/worksheets/sheet1.xml", `<worksheet xmlns="` + nsMain + `"><sheetData>` + sheetXML + `</sheetData></worksheet>`},
	}
	if sst != "" {
		es = append(es, entry{"xl/sharedStrings.xml", `<sst xmlns="` + nsMain + `">` + sst + `</sst>`})
	}
	return makeZip(t, es...)
}

func odsWith(t testing.TB, tableXML string) []byte {
	return makeZip(t,
		entry{"mimetype", mimeODS},
		entry{"content.xml", `<office:document-content xmlns:office="` + nsOffice + `" xmlns:table="` + nsTable + `" xmlns:text="` + nsText + `"><office:body><office:spreadsheet>` + tableXML + `</office:spreadsheet></office:body></office:document-content>`},
	)
}

func readBytes(b []byte, lim Limits) (*Workbook, error) {
	wb, _, err := Read(bytes.NewReader(b), int64(len(b)), &ReadOptions{Limits: lim})
	return wb, err
}

func TestMalicious(t *testing.T) {
	small := Limits{MaxCells: 100000, MaxEntrySize: 1 << 20}
	cases := []struct {
		name string
		data []byte
		want error
	}{
		{"xlsx huge row number", xlsxWith(t, `<row r="99999999"><c r="A99999999"><v>1</v></c></row>`, ""), ErrLimit},
		{"xlsx huge cell ref", xlsxWith(t, `<row><c r="XFE1"><v>1</v></c></row>`, ""), ErrLimit},
		{"xlsx negative sst index", xlsxWith(t, `<row><c t="s"><v>-1</v></c></row>`, `<si><t>a</t></si>`), ErrCorrupt},
		{"xlsx sst index out of range", xlsxWith(t, `<row><c t="s"><v>5</v></c></row>`, `<si><t>a</t></si>`), ErrCorrupt},
		{"xlsx sparse far cell", xlsxWith(t, `<row r="1000000"><c r="XFD1000000"><v>1</v></c></row>`, ""), ErrLimit},
		{"xlsx zip bomb entry", xlsxWith(t, `<row><c><v>1</v></c></row>`+strings.Repeat(" ", 2<<20), ""), ErrLimit},
		{"xlsx bad number", xlsxWith(t, `<row><c><v>abc</v></c></row>`, ""), ErrCorrupt},
		{"xlsx entity", xlsxWith(t, `<row><c t="inlineStr"><is><t>&xxe;</t></is></c></row>`, ""), ErrCorrupt},
		{"ods repeated non-empty rows x cols", odsWith(t, `<table:table table:name="A"><table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-repeated="16384" office:value-type="float" office:value="1"/></table:table-row></table:table>`), ErrLimit},
		{"ods repeat overflow", odsWith(t, `<table:table table:name="A"><table:table-row table:number-rows-repeated="9223372036854775807"><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row></table:table>`), ErrLimit},
		{"ods text:s bomb", odsWith(t, `<table:table table:name="A"><table:table-row><table:table-cell><text:p><text:s text:c="999999999"/></text:p></table:table-cell></table:table-row></table:table>`), ErrLimit},
		{"ods no tables", odsWith(t, ``), ErrCorrupt},
		{"not a zip", []byte("hello world, not a spreadsheet"), ErrUnsupported},
		{"ole xls", append(append([]byte{}, oleMagic...), make([]byte, 100)...), ErrUnsupported},
		{"plain zip", makeZip(t, entry{"a.txt", "x"}), ErrUnsupported},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			_, err := readBytes(c.data, small)
			if !errors.Is(err, c.want) {
				t.Fatalf("err = %v, want %v", err, c.want)
			}
		})
	}
}

func TestODSTrailingRepeatsAreFree(t *testing.T) {
	// common in LibreOffice files: trailing repeats fill up to the maximum row/column and must not allocate memory
	data := odsWith(t, `<table:table table:name="A">
<table:table-row><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell><table:table-cell table:number-columns-repeated="16383"/></table:table-row>
<table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>
</table:table>`)
	wb, err := readBytes(data, Limits{MaxCells: 10})
	if err != nil {
		t.Fatal(err)
	}
	if got := wb.Sheets[0].Strings(); len(got) != 1 || len(got[0]) != 1 || got[0][0] != "x" {
		t.Fatalf("got %q", got)
	}
}

func TestODSRepeatedValues(t *testing.T) {
	data := odsWith(t, `<table:table table:name="A">
<table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="3" office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell table:number-columns-repeated="2"/><table:table-cell office:value-type="string"><text:p>a<text:s/>b</text:p><office:annotation><text:p>comment</text:p></office:annotation></table:table-cell></table:table-row>
</table:table>`)
	wb, err := readBytes(data, Limits{})
	if err != nil {
		t.Fatal(err)
	}
	got := wb.Sheets[0].Strings()
	want := [][]string{{"7", "7", "7"}, {"7", "7", "7"}, {"", "", "a b"}}
	if len(got) != 3 || strings.Join(got[0], ",") != "7,7,7" || strings.Join(got[1], ",") != "7,7,7" || strings.Join(got[2], ",") != ",,a b" {
		t.Fatalf("got %q want %q", got, want)
	}
}

func TestDetect(t *testing.T) {
	for _, f := range []Format{FormatXLSX, FormatODS} {
		var buf bytes.Buffer
		if err := Write(&buf, sampleWorkbook(), f); err != nil {
			t.Fatal(err)
		}
		got, err := Detect(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
		if err != nil || got != f {
			t.Errorf("Detect = %v, %v; want %v", got, err, f)
		}
	}
}

func TestODSMimetypeFirstAndStored(t *testing.T) {
	var buf bytes.Buffer
	if err := Write(&buf, sampleWorkbook(), FormatODS); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	if zr.File[0].Name != "mimetype" || zr.File[0].Method != zip.Store {
		t.Fatalf("first entry = %s method %d", zr.File[0].Name, zr.File[0].Method)
	}
	// required for Excel to open ODS (tested 2026-11): no data descriptor, a valid date, ReaderVersion 20
	for _, f := range zr.File {
		if f.Flags&0x8 != 0 || f.ModifiedDate == 0 || f.ReaderVersion != 20 { //nolint:staticcheck
			t.Errorf("%s: flags=%x date=%d readerVersion=%d", f.Name, f.Flags, f.ModifiedDate, f.ReaderVersion) //nolint:staticcheck
		}
	}
	// the mimetype string must start right after the first 38 bytes (magic number convention)
	if !bytes.Contains(buf.Bytes()[:100], []byte(mimeODS)) {
		t.Fatal("mimetype not at fixed offset")
	}
}

func TestCellRef(t *testing.T) {
	for _, c := range []struct {
		r, c int
		s    string
	}{{0, 0, "A1"}, {9, 25, "Z10"}, {0, 26, "AA1"}, {1048575, 16383, "XFD1048576"}} {
		if got := CellRef(c.r, c.c); got != c.s {
			t.Errorf("CellRef(%d,%d) = %s", c.r, c.c, got)
		}
		r, cc, ok := parseCellRef(c.s)
		if !ok || r != c.r || cc != c.c {
			t.Errorf("parse(%s) = %d,%d,%v", c.s, r, cc, ok)
		}
	}
	for _, s := range []string{"", "A", "1", "A0", "a1", "AAAA1", "A12345678"} {
		if _, _, ok := parseCellRef(s); ok {
			t.Errorf("parse(%q) should fail", s)
		}
	}
}

func FuzzRead(f *testing.F) {
	for _, fm := range []Format{FormatXLSX, FormatODS} {
		var buf bytes.Buffer
		if err := Write(&buf, sampleWorkbook(), fm); err != nil {
			f.Fatal(err)
		}
		f.Add(buf.Bytes())
	}
	f.Add(odsWith(f, `<table:table table:name="A"><table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="2" office:value-type="date" office:date-value="2026-01-02T03:04:05"/></table:table-row></table:table>`))
	f.Add(xlsxWith(f, `<row r="2"><c r="B2" t="s"><v>0</v></c><c t="b"><v>1</v></c></row>`, `<si><r><t>a</t></r><rPh><t>x</t></rPh></si>`))
	f.Fuzz(func(t *testing.T, data []byte) {
		o := ReadOptions{Limits: Limits{MaxCells: 200000, MaxEntrySize: 4 << 20, MaxTotalSize: 8 << 20}.withDefaults(), Location: time.UTC}
		// call the internal read directly (no recover), so the fuzzer catches any panic
		wb, _, err := read(bytes.NewReader(data), int64(len(data)), &o)
		if err == nil && wb != nil {
			for _, fm := range []Format{FormatXLSX, FormatODS} {
				// a file that can be read must not panic when written back (it may return an error for an invalid name)
				_ = Write(&bytes.Buffer{}, wb, fm)
			}
		}
	})
}
