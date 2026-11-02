package sheetio

import (
	"bytes"
	"errors"
	"hash/crc32"
	"math"
	"strconv"
	"strings"
	"testing"
)

func TestCRC32Combine(t *testing.T) {
	parts := []string{"", "a", "hello ", "世界", strings.Repeat("xyz", 10000)}
	for _, p1 := range parts {
		for _, p2 := range parts {
			want := crc32.ChecksumIEEE([]byte(p1 + p2))
			got := crc32Combine(crc32.ChecksumIEEE([]byte(p1)), crc32.ChecksumIEEE([]byte(p2)), int64(len(p2)))
			if got != want {
				t.Fatalf("combine(%q,%q) = %x want %x", p1[:min(len(p1), 5)], p2[:min(len(p2), 5)], got, want)
			}
		}
	}
}

func TestStreamWriteReadRows(t *testing.T) {
	for _, f := range []Format{FormatXLSX, FormatODS} {
		t.Run(f.String(), func(t *testing.T) {
			var buf bytes.Buffer
			bold := Style{Bold: true}
			sw, err := NewStreamWriter(&buf, f, []Style{bold}, []SheetDef{
				{Name: "大量", ColWidths: []float64{10, 20}},
				{Name: "合併", Merges: []Range{{0, 0, 0, 3}, {5, 1, 7, 2}}}, // the second merge extends past the written rows
				{Name: "沒寫到"},
			})
			if err != nil {
				t.Fatal(err)
			}
			const n = 5000
			for i := 0; i < n; i++ {
				if i == 100 {
					if err := sw.WriteRow(); err != nil { // empty row
						t.Fatal(err)
					}
					continue
				}
				if err := sw.WriteRow(Int(int64(i)), Str("列"+strconv.Itoa(i)).WithStyle(1)); err != nil {
					t.Fatal(err)
				}
			}
			// an invalid row is rejected, but the stream can continue
			if err := sw.WriteRow(Num(math.Inf(1))); !errors.Is(err, ErrInvalid) {
				t.Fatalf("want ErrInvalid, got %v", err)
			}
			if err := sw.NextSheet(); err != nil {
				t.Fatal(err)
			}
			if err := sw.WriteRow(Str("標題")); err != nil {
				t.Fatal(err)
			}
			if err := sw.WriteRow(Str("a"), Str("b")); err != nil {
				t.Fatal(err)
			}
			if err := sw.Close(); err != nil {
				t.Fatal(err)
			}
			if err := sw.WriteRow(Str("x")); !errors.Is(err, ErrClosed) {
				t.Fatalf("after close: %v", err)
			}

			// read back as a stream
			counts := map[string]int{}
			last := -1
			_, err = ReadRows(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil, func(sheet string, r int, cells []Cell) error {
				counts[sheet]++
				if sheet == "大量" {
					if r <= last {
						t.Fatalf("rows not ascending: %d after %d", r, last)
					}
					last = r
					if r == 100 {
						t.Fatal("blank row delivered")
					}
					if got := cells[0].String(); got != strconv.Itoa(r) || cells[1].Str != "列"+strconv.Itoa(r) {
						t.Fatalf("row %d = %q", r, got)
					}
				}
				return nil
			})
			if err != nil {
				t.Fatal(err)
			}
			if counts["大量"] != n-1 || counts["合併"] != 2 || counts["沒寫到"] != 0 {
				t.Fatalf("counts = %v", counts)
			}

			// full read: empty sheets must still exist
			wb, _, err := Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
			if err != nil {
				t.Fatal(err)
			}
			if len(wb.Sheets) != 3 || len(wb.Sheets[0].Rows) != n || wb.Sheets[0].Rows[100] != nil {
				t.Fatalf("sheets=%d rows=%d", len(wb.Sheets), len(wb.Sheets[0].Rows))
			}
		})
	}
}

func TestReadRowsSkipAndStop(t *testing.T) {
	for _, f := range []Format{FormatXLSX, FormatODS} {
		var buf bytes.Buffer
		if err := Write(&buf, sampleWorkbook(), f); err != nil {
			t.Fatal(err)
		}
		// ErrSkipSheet: read only the first row of each sheet
		seen := map[string]int{}
		_, err := ReadRows(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil, func(sheet string, r int, cells []Cell) error {
			seen[sheet]++
			return ErrSkipSheet
		})
		if err != nil || seen["員工清單"] != 1 || seen["Summary"] != 1 {
			t.Fatalf("%v: skip: %v %v", f, seen, err)
		}
		// other errors are returned as is
		stop := errors.New("stop")
		_, err = ReadRows(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil, func(string, int, []Cell) error { return stop })
		if err != stop {
			t.Fatalf("%v: want stop, got %v", f, err)
		}
	}
}

func TestReadRowsODSRepeatedRows(t *testing.T) {
	data := odsWith(t, `<table:table table:name="A"><table:table-row table:number-rows-repeated="3"><table:table-cell office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row></table:table>`)
	var rows []int
	_, err := ReadRows(bytes.NewReader(data), int64(len(data)), nil, func(_ string, r int, c []Cell) error {
		rows = append(rows, r)
		return nil
	})
	if err != nil || len(rows) != 3 || rows[2] != 2 {
		t.Fatalf("rows=%v err=%v", rows, err)
	}
	// MaxCells also caps what is delivered while streaming (so repeats cannot burn CPU)
	data = odsWith(t, `<table:table table:name="A"><table:table-row table:number-rows-repeated="1000000"><table:table-cell table:number-columns-repeated="1000" office:value-type="float" office:value="1"/></table:table-row></table:table>`)
	_, err = ReadRows(bytes.NewReader(data), int64(len(data)), &ReadOptions{Limits: Limits{MaxCells: 50000}}, func(string, int, []Cell) error { return nil })
	if !errors.Is(err, ErrLimit) {
		t.Fatalf("want ErrLimit, got %v", err)
	}
}

func TestStreamWriterDefs(t *testing.T) {
	if _, err := NewStreamWriter(&bytes.Buffer{}, FormatXLSX, nil, nil); !errors.Is(err, ErrInvalid) {
		t.Fatalf("no sheets: %v", err)
	}
	sw, err := NewStreamWriter(&bytes.Buffer{}, FormatODS, nil, []SheetDef{{Name: "A"}})
	if err != nil {
		t.Fatal(err)
	}
	if err := sw.WriteRow(Str("a")); err != nil {
		t.Fatal(err)
	}
	if err := sw.NextSheet(); !errors.Is(err, ErrInvalid) {
		t.Fatalf("next beyond defs: %v", err)
	}
	if err := sw.WriteRow(Str("b").WithStyle(3)); !errors.Is(err, ErrInvalid) {
		t.Fatalf("bad style: %v", err)
	}
	if err := sw.Close(); err != nil {
		t.Fatal(err)
	}
}
