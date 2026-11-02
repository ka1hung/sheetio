# sheetio

[![Go Reference](https://pkg.go.dev/badge/github.com/ka1hung/sheetio.svg)](https://pkg.go.dev/github.com/ka1hung/sheetio)
[![CI](https://github.com/ka1hung/sheetio/actions/workflows/ci.yml/badge.svg)](https://github.com/ka1hung/sheetio/actions/workflows/ci.yml)

English | [繁體中文](README.zh-TW.md)

Read and write **xlsx** and **ods** spreadsheets in Go, using **only the standard library**.

sheetio is built for the everyday job of exporting and importing tabular data: reports, data
dumps, bulk uploads. It does not try to be a full spreadsheet engine.

- Zero third-party dependencies (`archive/zip` + `encoding/xml`)
- xlsx and ods, auto-detected from file content (not the file name)
- Strings, numbers, booleans, dates; bold, fill color, borders, alignment, wrap, number formats
- Column widths, row heights, merged cells, freeze panes, auto filter
- Streaming writer: memory use does not grow with the number of rows, no temp files
- Streaming reader: row-by-row callback, no temp files
- Configurable limits for untrusted uploads

```
go get github.com/ka1hung/sheetio
```

Requires Go 1.21 or later.

## Contents

- [Write a file](#write-a-file)
- [Read a file](#read-a-file)
- [Streaming write](#streaming-write)
- [Streaming read](#streaming-read)
- [Gin examples](#gin-examples)
- [Errors](#errors)
- [Reference](#reference)
- [Performance](#performance)
- [Limitations](#limitations)

## Write a file

```go
package main

import (
	"log"
	"os"
	"time"

	"github.com/ka1hung/sheetio"
)

func main() {
	wb := sheetio.New()

	// Styles are registered once and referenced by ID.
	head := wb.AddStyle(sheetio.Style{Bold: true, FillColor: "D9E1F2", Border: true, Align: sheetio.AlignCenter})
	money := wb.AddStyle(sheetio.Style{NumFmt: sheetio.NumFmtThousands})

	s := wb.AddSheet("Sales")
	s.SetColWidth(0, 14) // column A, in characters
	s.SetColWidth(1, 20)
	s.SetRowHeight(0, 24) // row 1, in points

	s.AddRow(
		sheetio.Str("Product").WithStyle(head),
		sheetio.Str("Date").WithStyle(head),
		sheetio.Str("Amount").WithStyle(head),
	)
	s.AddRow(sheetio.Str("A-100"), sheetio.Date(time.Date(2026, 11, 1, 9, 30, 0, 0, time.Local)), sheetio.Num(12345.6).WithStyle(money))
	s.AddRow(sheetio.Str("B-200"), sheetio.Date(time.Date(2026, 11, 2, 0, 0, 0, 0, time.Local)), sheetio.Int(980).WithStyle(money))

	s.Freeze(1, 0)           // keep the header row visible
	s.SetAutoFilter(0, 0, 2) // filter buttons on A1:C1, range extends to the last row

	// A second sheet with a merged title.
	t := wb.AddSheet("Summary")
	t.AddRow(sheetio.Str("November 2026").WithStyle(head))
	t.Merge(0, 0, 0, 2) // A1:C1

	f, err := os.Create("sales.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()
	if err := sheetio.Write(f, wb, sheetio.FormatXLSX); err != nil { // or sheetio.FormatODS
		log.Fatal(err)
	}
}
```

All row and column indexes are **0-based** (`Merge(0, 0, 0, 2)` is `A1:C1`).

`Write` accepts any `io.Writer`, so the same call writes to a file, a `bytes.Buffer`, or an HTTP
response:

```go
var buf bytes.Buffer
if err := sheetio.Write(&buf, wb, sheetio.FormatODS); err != nil {
	return err
}
// buf.Bytes() is a complete .ods file
```

`Write` validates the workbook first and returns `ErrInvalid` (nothing is written) for things like
duplicate sheet names, invalid UTF-8, NaN, or out-of-range dates. Call `wb.Validate()` to check
without writing.

## Read a file

`Read` takes an `io.ReaderAt` and its size. `*os.File`, `*bytes.Reader` and `multipart.File` all
work. The format is detected from the content.

```go
f, err := os.Open("sales.xlsx")
if err != nil {
	return err
}
defer f.Close()
st, err := f.Stat()
if err != nil {
	return err
}

wb, format, err := sheetio.Read(f, st.Size(), nil) // nil = default options
if err != nil {
	return err
}
fmt.Println("format:", format) // xlsx

for _, sh := range wb.Sheets {
	fmt.Println("sheet:", sh.Name)

	// Easiest: everything as strings. Dates become "2006-01-02" or "2006-01-02 15:04:05",
	// booleans become "TRUE" / "FALSE". Trailing empty cells and rows are trimmed.
	for i, row := range sh.Strings() {
		fmt.Println(i, row)
	}
}

// Typed access: Cell(row, col) returns an empty cell when out of range.
sh := wb.Sheet("sales") // case-insensitive, nil if not found
c := sh.Cell(1, 2)
switch c.Kind {
case sheetio.KindNumber:
	fmt.Println("number", c.Num)
case sheetio.KindDate:
	fmt.Println("date", c.Time)
case sheetio.KindString:
	fmt.Println("string", c.Str)
case sheetio.KindBool:
	fmt.Println("bool", c.Bool)
case sheetio.KindEmpty:
	fmt.Println("empty")
}
```

Options:

```go
wb, _, err := sheetio.Read(f, size, &sheetio.ReadOptions{
	// Dates have no time zone in the file; this is the zone they are placed in (default time.Local).
	Location: time.UTC,
	// Zero fields fall back to DefaultLimits().
	Limits: sheetio.Limits{MaxFileSize: 20 << 20, MaxSheets: 10, MaxRows: 100000, MaxCols: 200},
})
if err != nil {
	return err
}
```

`Read` keeps the whole workbook in memory. For large or sparse files use
[`ReadRows`](#streaming-read).

## Streaming write

`NewStreamWriter` writes rows as they are produced, so millions of rows can be exported with a
few MB of memory. xlsx is written straight to the output. ods keeps only the compressed content
in memory, because the format needs it before the end of the file.

Sheet settings (widths, heights, merges, freeze, filter) are declared up front in `SheetDef`.
Sheets are written in order: finish one, call `NextSheet`, start the next.

```go
sw, err := sheetio.NewStreamWriter(w, sheetio.FormatXLSX,
	[]sheetio.Style{{Bold: true}, {NumFmt: sheetio.NumFmtDateTime}}, // style IDs 1 and 2
	[]sheetio.SheetDef{
		{
			Name:       "Logs",
			ColWidths:  []float64{20, 12, 40},
			FreezeRows: 1,
			AutoFilter: &sheetio.Range{Row1: 0, Col1: 0, Row2: 0, Col2: 2}, // header only; extends to the last row
		},
		{Name: "Totals"},
	})
if err != nil {
	return err
}

bold := 1
sw.WriteRow(sheetio.Str("Time").WithStyle(bold), sheetio.Str("Level").WithStyle(bold), sheetio.Str("Message").WithStyle(bold))

row := make([]sheetio.Cell, 3) // the slice can be reused between calls
for rows.Next() {              // e.g. *sql.Rows
	var t time.Time
	var level, msg string
	if err := rows.Scan(&t, &level, &msg); err != nil {
		return err
	}
	row[0] = sheetio.Date(t).WithStyle(2)
	row[1] = sheetio.Str(level)
	row[2] = sheetio.Str(msg)
	if err := sw.WriteRow(row...); err != nil {
		if errors.Is(err, sheetio.ErrInvalid) {
			continue // this row was rejected (e.g. NaN); the stream is still usable
		}
		return err
	}
}

if err := sw.NextSheet(); err != nil { // switch to "Totals"
	return err
}
sw.WriteRow(sheetio.Str("count"), sheetio.Int(42))

return sw.Close() // always call Close; it finishes the file
```

`Write(w, wb, f)` uses the same writer internally.

## Streaming read

`ReadRows` parses the file row by row and calls your function for each non-empty row. Nothing
except the shared string table is kept in memory.

```go
f, err := os.Open("logs.xlsx")
if err != nil {
	return err
}
defer f.Close()
st, _ := f.Stat()

var count int
_, err = sheetio.ReadRows(f, st.Size(), nil, func(sheet string, row int, cells []sheetio.Cell) error {
	if sheet != "Logs" {
		return sheetio.ErrSkipSheet // skip the rest of this sheet
	}
	if row == 0 {
		return nil // header
	}
	// cells are reused after the callback returns: copy anything you keep.
	level := ""
	if len(cells) > 1 {
		level = cells[1].String()
	}
	if level == "ERROR" {
		count++
	}
	if count > 1000 {
		return errStop // any other error stops reading and is returned as is
	}
	return nil
})
if err != nil && !errors.Is(err, errStop) {
	return err
}
```

- `row` is the 0-based row number. Empty rows are skipped, so numbers can have gaps.
- Trailing empty cells are already trimmed (`len(cells)` may differ between rows).
- Completely empty sheets produce no callback. Use `Read` if you need the list of sheets.
- Default limits are wider than `Read`'s (100 million cells, 2 GB per entry), since memory use is
  flat. Set your own `Limits` for user uploads.

## Gin examples

### Download (small or medium data)

```go
r.GET("/report", func(c *gin.Context) {
	f := sheetio.FormatXLSX
	if c.Query("format") == "ods" {
		f = sheetio.FormatODS
	}

	wb := sheetio.New()
	s := wb.AddSheet("Report")
	s.AddRow(sheetio.Str("ID"), sheetio.Str("Name"))
	for _, u := range loadUsers() {
		s.AddRow(sheetio.Int(u.ID), sheetio.Str(u.Name))
	}

	var buf bytes.Buffer // build first, so errors can still return a 500
	if err := sheetio.Write(&buf, wb, f); err != nil {
		c.String(http.StatusInternalServerError, err.Error())
		return
	}
	name := "report_" + time.Now().Format("20060102") + f.Ext()
	c.Header("Content-Disposition", "attachment; filename*=UTF-8''"+url.PathEscape(name))
	c.Data(http.StatusOK, f.ContentType(), buf.Bytes())
})
```

### Download (large data, streamed)

```go
r.GET("/logs.xlsx", func(c *gin.Context) {
	rows, err := db.QueryContext(c, "SELECT time, level, message FROM logs ORDER BY time")
	if err != nil {
		c.String(http.StatusInternalServerError, err.Error())
		return
	}
	defer rows.Close()

	c.Header("Content-Type", sheetio.FormatXLSX.ContentType())
	c.Header("Content-Disposition", `attachment; filename="logs.xlsx"`)

	sw, err := sheetio.NewStreamWriter(c.Writer, sheetio.FormatXLSX, nil,
		[]sheetio.SheetDef{{Name: "Logs", FreezeRows: 1, AutoFilter: &sheetio.Range{Col2: 2}}})
	if err != nil {
		c.String(http.StatusInternalServerError, err.Error())
		return
	}
	sw.WriteRow(sheetio.Str("Time"), sheetio.Str("Level"), sheetio.Str("Message"))
	for rows.Next() {
		var t time.Time
		var level, msg string
		if err := rows.Scan(&t, &level, &msg); err != nil {
			log.Println(err) // headers are already sent; the download will be incomplete
			return
		}
		sw.WriteRow(sheetio.Date(t), sheetio.Str(level), sheetio.Str(msg))
	}
	if err := sw.Close(); err != nil {
		log.Println(err)
	}
})
```

### Upload

```go
r.POST("/import", func(c *gin.Context) {
	c.Request.Body = http.MaxBytesReader(c.Writer, c.Request.Body, 20<<20)
	fh, err := c.FormFile("file")
	if err != nil {
		c.String(http.StatusBadRequest, "missing file")
		return
	}
	f, err := fh.Open() // multipart.File implements io.ReaderAt
	if err != nil {
		c.String(http.StatusBadRequest, "cannot open file")
		return
	}
	defer f.Close()

	opt := &sheetio.ReadOptions{Limits: sheetio.Limits{MaxFileSize: 20 << 20, MaxRows: 100000, MaxCols: 50}}
	var users []User
	_, err = sheetio.ReadRows(f, fh.Size, opt, func(sheet string, row int, cells []sheetio.Cell) error {
		if row == 0 {
			return nil // header
		}
		if len(cells) < 2 {
			return fmt.Errorf("row %d: expected 2 columns", row+1)
		}
		users = append(users, User{ID: int64(cells[0].Num), Name: cells[1].String()})
		return nil
	})
	switch {
	case errors.Is(err, sheetio.ErrUnsupported):
		c.String(http.StatusBadRequest, "only .xlsx and .ods are accepted")
		return
	case errors.Is(err, sheetio.ErrLimit), errors.Is(err, sheetio.ErrCorrupt):
		c.String(http.StatusBadRequest, "file too large or damaged")
		return
	case err != nil:
		c.String(http.StatusBadRequest, err.Error())
		return
	}
	c.JSON(http.StatusOK, gin.H{"imported": len(users)})
})
```

## Errors

| Error | When |
|---|---|
| `ErrInvalid` | Writing: the workbook or a row is not valid (bad sheet name, NaN, date out of range, …) |
| `ErrClosed` | Writing: `StreamWriter` used after `Close` |
| `ErrUnsupported` | Reading: not xlsx/ods (including legacy .xls and encrypted files) |
| `ErrCorrupt` | Reading: the file is damaged or does not follow the format |
| `ErrLimit` | Reading: a configured limit was exceeded |
| `ErrSkipSheet` | Returned from a `ReadRows` callback to skip the rest of the sheet |

All errors are wrapped; use `errors.Is`.

## Reference

### Cells

| Constructor | Notes |
|---|---|
| `Str(s)` | Always a text cell; values starting with `=` are not treated as formulas |
| `Num(f)`, `Int(i)` | Stored as float64. Integers above 2^53 lose precision: use `Str` for IDs, card numbers, etc. |
| `Bool(b)` | |
| `Date(t)` | Wall-clock time, no time zone. Range 1900-03-01 to 9999-12-31. Formatted as a date or date-time automatically |
| `Empty()` | Blank cell, useful with `WithStyle` for a styled empty cell |
| `.WithStyle(id)` | Applies a style returned by `AddStyle` |

### Style

```go
sheetio.Style{
	Bold:      true,
	FillColor: "FFF2CC", // RRGGBB
	Border:    true,     // thin border on all sides
	Align:     sheetio.AlignCenter, // AlignLeft, AlignCenter, AlignRight
	Wrap:      true,
	NumFmt:    sheetio.NumFmtDecimal2,
}
```

Number formats: `NumFmtGeneral`, `NumFmtInteger` (`0`), `NumFmtDecimal2` (`0.00`),
`NumFmtThousands` (`#,##0`), `NumFmtDate` (`yyyy-mm-dd`), `NumFmtDateTime` (`yyyy-mm-dd hh:mm:ss`).

### Sheet layout

| Method / field | Notes |
|---|---|
| `SetColWidth(col, width)` | Width in characters (0–255) |
| `SetRowHeight(row, height)` | Height in points (1–409). Applies to rows that are written |
| `Merge(r1, c1, r2, c2)` | Merged ranges must not overlap |
| `Freeze(rows, cols)` | Freeze the top rows and left columns |
| `SetAutoFilter(headerRow, firstCol, lastCol)` | Filter buttons on the header row; the range extends to the last data row |

`StreamWriter` takes the same settings in `SheetDef` (`ColWidths`, `RowHeights`, `Merges`,
`FreezeRows`, `FreezeCols`, `AutoFilter`).

### Default read limits

| Field | `Read` | `ReadRows` |
|---|---|---|
| `MaxFileSize` (compressed) | 50 MB | 50 MB |
| `MaxEntries` (zip entries) | 10,000 | 10,000 |
| `MaxEntrySize` (one entry, uncompressed) | 256 MB | 2 GB |
| `MaxTotalSize` (all entries, uncompressed) | 512 MB | 4 GB |
| `MaxSheets` | 256 | 256 |
| `MaxRows` / `MaxCols` | 1,048,576 / 16,384 | same |
| `MaxCells` | 5,000,000 | 100,000,000 |
| `MaxCellChars` | 32,767 | same |

### Helpers

`ColName(2)` → `"C"`, `CellRef(0, 2)` → `"C1"`, `Detect(r, size)` returns the format without
parsing, `ValidateSheetName(name)`, `Format.Ext()` / `Format.ContentType()` / `Format.String()`.

## Performance

10 columns per row (integers, strings, dates, decimals, booleans), measured on a 2-core cloud
container. Peak memory is the whole process RSS.

| Operation | 100k rows | 1M rows | Peak memory (1M rows) |
|---|---|---|---|
| Streaming write, xlsx | 1.0 s | 9.7 s | 8 MB |
| Streaming read, xlsx (`ReadRows`) | 2.4 s | 23.7 s | 11 MB |
| Streaming write, ods | 1.2 s | 12.0 s | 104 MB (compressed content.xml) |
| Streaming read, ods (`ReadRows`) | 4.6 s | 47 s | 9 MB |
| Full read, xlsx (`Read`) | 3.1 s | — | 151 MB (100k rows) |

`Read` builds a `Cell` for every cell, including empty ones inside a row. For large or sparse
sheets, `ReadRows` is faster and uses much less memory.

## Limitations

- **Formulas** are not evaluated. Reading returns the value cached by the program that saved the
  file. Files from tools that do not calculate formulas will have empty formula cells.
- **Numbers** are float64. Write IDs and other digit strings with `Str` to keep leading zeros.
- **Dates** have no time zone: the wall-clock time is written, and dates are read into
  `ReadOptions.Location`.
- **Styles, widths, heights, merges, freeze panes and filters are write-only.** Reading returns
  values only (the top-left value of a merged range).
- **Not supported:** charts, images, pivot tables, conditional formatting, comments, data
  validation, legacy `.xls`, encrypted files.
- **ods column widths** are converted from Excel character units to centimetres, so they look
  slightly different from xlsx.
- **ods freeze panes** are stored in `settings.xml`. Excel does not apply freeze panes when it
  opens an ods file (filters and row heights work).
- A `StreamWriter` writes one sheet at a time; sheets cannot be interleaved.

## License

[MIT](LICENSE)
