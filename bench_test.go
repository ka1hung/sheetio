package sheetio

import (
	"bytes"
	"strconv"
	"testing"
	"time"
)

func bigWorkbook(rows int) *Workbook {
	wb := New()
	s := wb.AddSheet("Data")
	d := time.Date(2026, 1, 1, 8, 0, 0, 0, time.UTC)
	for i := 0; i < rows; i++ {
		s.AddRow(Int(int64(i)), Str("Device-"+strconv.Itoa(i)), Num(float64(i)*1.5), Date(d.Add(time.Duration(i)*time.Minute)),
			Bool(i%2 == 0), Str("status note"), Num(220.1), Num(3.3), Str("OK"), Int(int64(i*7)))
	}
	return wb
}

func benchRT(b *testing.B, f Format) {
	wb := bigWorkbook(100000)
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		var buf bytes.Buffer
		if err := Write(&buf, wb, f); err != nil {
			b.Fatal(err)
		}
		b.ReportMetric(float64(buf.Len())/1e6, "MB")
		if _, _, err := Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkRoundTrip100kXLSX(b *testing.B) { benchRT(b, FormatXLSX) }
func BenchmarkRoundTrip100kODS(b *testing.B)  { benchRT(b, FormatODS) }
