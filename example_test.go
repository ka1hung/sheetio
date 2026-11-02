package sheetio_test

import (
	"bytes"
	"fmt"
	"time"

	"github.com/ka1hung/sheetio"
)

func Example() {
	// ---- write ----
	wb := sheetio.New()
	head := wb.AddStyle(sheetio.Style{Bold: true, FillColor: "D9E1F2", Border: true, Align: sheetio.AlignCenter})

	s := wb.AddSheet("Access log")
	s.SetColWidth(1, 20)
	s.AddRow(sheetio.Str("Card").WithStyle(head), sheetio.Str("Time").WithStyle(head), sheetio.Str("Door").WithStyle(head))
	s.AddRow(sheetio.Str("00123"), sheetio.Date(time.Date(2026, 11, 1, 8, 59, 3, 0, time.Local)), sheetio.Str("Main entrance"))
	s.Freeze(1, 0)
	s.SetAutoFilter(0, 0, 2)

	var buf bytes.Buffer
	if err := sheetio.Write(&buf, wb, sheetio.FormatODS); err != nil {
		panic(err)
	}

	// ---- read (format detected from the content) ----
	in, format, err := sheetio.Read(bytes.NewReader(buf.Bytes()), int64(buf.Len()), nil)
	if err != nil {
		panic(err)
	}
	fmt.Println(format, in.Sheets[0].Name)
	for _, row := range in.Sheets[0].Strings() {
		fmt.Printf("%q\n", row)
	}
	// Output:
	// ods Access log
	// ["Card" "Time" "Door"]
	// ["00123" "2026-11-01 08:59:03" "Main entrance"]
}
