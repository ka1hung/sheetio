# sheetio

[![Go Reference](https://pkg.go.dev/badge/github.com/ka1hung/sheetio.svg)](https://pkg.go.dev/github.com/ka1hung/sheetio)
[![CI](https://github.com/ka1hung/sheetio/actions/workflows/ci.yml/badge.svg)](https://github.com/ka1hung/sheetio/actions/workflows/ci.yml)

[English](README.md) | 繁體中文

用 Go 讀寫 **xlsx** 與 **ods** 試算表，**只依賴標準庫**。

sheetio 專注在日常的資料表格匯出／匯入：報表、資料下載、批次上傳。它不是完整的試算表引擎。

- 沒有任何第三方相依（`archive/zip` + `encoding/xml`）
- 支援 xlsx 與 ods，依檔案內容判斷格式（不看副檔名）
- 字串、數字、布林、日期；粗體、底色、框線、對齊、自動換行、數字格式
- 欄寬、列高、合併儲存格、凍結窗格、自動篩選
- 串流寫出：記憶體用量不隨列數增加，不使用暫存檔
- 串流讀取：逐列 callback，不使用暫存檔
- 讀取時可設定資源上限，適合處理使用者上傳的檔案

```
go get github.com/ka1hung/sheetio
```

需要 Go 1.21 以上。

## 目錄

- [寫出檔案](#寫出檔案)
- [讀取檔案](#讀取檔案)
- [串流寫出](#串流寫出)
- [串流讀取](#串流讀取)
- [Gin 範例](#gin-範例)
- [錯誤](#錯誤)
- [參考](#參考)
- [效能](#效能)
- [限制](#限制)

## 寫出檔案

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

	// 樣式先註冊一次，之後用 ID 引用
	head := wb.AddStyle(sheetio.Style{Bold: true, FillColor: "D9E1F2", Border: true, Align: sheetio.AlignCenter})
	money := wb.AddStyle(sheetio.Style{NumFmt: sheetio.NumFmtThousands})

	s := wb.AddSheet("銷售")
	s.SetColWidth(0, 14) // A 欄，單位是字元數
	s.SetColWidth(1, 20)
	s.SetRowHeight(0, 24) // 第 1 列，單位是點

	s.AddRow(
		sheetio.Str("產品").WithStyle(head),
		sheetio.Str("日期").WithStyle(head),
		sheetio.Str("金額").WithStyle(head),
	)
	s.AddRow(sheetio.Str("A-100"), sheetio.Date(time.Date(2026, 11, 1, 9, 30, 0, 0, time.Local)), sheetio.Num(12345.6).WithStyle(money))
	s.AddRow(sheetio.Str("B-200"), sheetio.Date(time.Date(2026, 11, 2, 0, 0, 0, 0, time.Local)), sheetio.Int(980).WithStyle(money))

	s.Freeze(1, 0)           // 凍結標題列
	s.SetAutoFilter(0, 0, 2) // A1:C1 加上篩選按鈕，範圍自動延伸到最後一列

	// 第二張表，標題合併儲存格
	t := wb.AddSheet("摘要")
	t.AddRow(sheetio.Str("2026 年 11 月").WithStyle(head))
	t.Merge(0, 0, 0, 2) // A1:C1

	f, err := os.Create("sales.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()
	if err := sheetio.Write(f, wb, sheetio.FormatXLSX); err != nil { // 或 sheetio.FormatODS
		log.Fatal(err)
	}
}
```

所有列號、欄號都是 **0 起算**（`Merge(0, 0, 0, 2)` 就是 `A1:C1`）。

`Write` 接受任何 `io.Writer`，所以同一個呼叫可以寫到檔案、`bytes.Buffer` 或 HTTP 回應：

```go
var buf bytes.Buffer
if err := sheetio.Write(&buf, wb, sheetio.FormatODS); err != nil {
	return err
}
// buf.Bytes() 就是完整的 .ods 檔
```

`Write` 會先檢查活頁簿，遇到工作表名稱重複、非法 UTF-8、NaN、日期超出範圍等情況會回傳
`ErrInvalid`，而且不會寫出任何內容。只想檢查不寫出時可以呼叫 `wb.Validate()`。

## 讀取檔案

`Read` 接受 `io.ReaderAt` 和檔案大小，`*os.File`、`*bytes.Reader`、`multipart.File` 都可以直接用。
格式依內容判斷。

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

wb, format, err := sheetio.Read(f, st.Size(), nil) // nil = 預設選項
if err != nil {
	return err
}
fmt.Println("格式:", format) // xlsx

for _, sh := range wb.Sheets {
	fmt.Println("工作表:", sh.Name)

	// 最簡單：全部轉字串。日期會變成 "2006-01-02" 或 "2006-01-02 15:04:05"，
	// 布林是 "TRUE" / "FALSE"。每列尾端的空白格與尾端空白列都已去除。
	for i, row := range sh.Strings() {
		fmt.Println(i, row)
	}
}

// 依型別取值：Cell(row, col) 超出範圍時回傳空白格
sh := wb.Sheet("銷售") // 名稱不分大小寫，找不到回傳 nil
c := sh.Cell(1, 2)
switch c.Kind {
case sheetio.KindNumber:
	fmt.Println("數字", c.Num)
case sheetio.KindDate:
	fmt.Println("日期", c.Time)
case sheetio.KindString:
	fmt.Println("字串", c.Str)
case sheetio.KindBool:
	fmt.Println("布林", c.Bool)
case sheetio.KindEmpty:
	fmt.Println("空白")
}
```

選項：

```go
wb, _, err := sheetio.Read(f, size, &sheetio.ReadOptions{
	// 檔案裡的日期沒有時區，讀出時放在這個時區（預設 time.Local）
	Location: time.UTC,
	// 欄位為 0 時使用 DefaultLimits() 的值
	Limits: sheetio.Limits{MaxFileSize: 20 << 20, MaxSheets: 10, MaxRows: 100000, MaxCols: 200},
})
if err != nil {
	return err
}
```

`Read` 會把整份活頁簿放進記憶體。檔案很大或很稀疏時，請改用 [`ReadRows`](#串流讀取)。

## 串流寫出

`NewStreamWriter` 邊產生邊寫出，幾百萬列也只用幾 MB 記憶體。xlsx 直接寫進輸出；
ods 因為格式要求，只會在記憶體保留壓縮後的內容。

工作表設定（欄寬、列高、合併、凍結、篩選）要事先寫在 `SheetDef`。
工作表依序寫出：寫完一張，呼叫 `NextSheet`，再寫下一張。

```go
sw, err := sheetio.NewStreamWriter(w, sheetio.FormatXLSX,
	[]sheetio.Style{{Bold: true}, {NumFmt: sheetio.NumFmtDateTime}}, // 樣式 ID 1、2
	[]sheetio.SheetDef{
		{
			Name:       "紀錄",
			ColWidths:  []float64{20, 12, 40},
			FreezeRows: 1,
			AutoFilter: &sheetio.Range{Row1: 0, Col1: 0, Row2: 0, Col2: 2}, // 只給標題列，自動延伸到最後一列
		},
		{Name: "統計"},
	})
if err != nil {
	return err
}

bold := 1
sw.WriteRow(sheetio.Str("時間").WithStyle(bold), sheetio.Str("等級").WithStyle(bold), sheetio.Str("訊息").WithStyle(bold))

row := make([]sheetio.Cell, 3) // slice 可以重複使用
for rows.Next() {              // 例如 *sql.Rows
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
			continue // 這一列不合法（例如 NaN）沒有寫出，串流仍可繼續
		}
		return err
	}
}

if err := sw.NextSheet(); err != nil { // 換到「統計」
	return err
}
sw.WriteRow(sheetio.Str("筆數"), sheetio.Int(42))

return sw.Close() // 一定要呼叫，才會完成檔案
```

`Write(w, wb, f)` 內部也是走同一個寫出器。

## 串流讀取

`ReadRows` 逐列解析檔案，每讀到一列非空白資料就呼叫一次你的函式。除了共用字串表以外，
不會把資料留在記憶體。

```go
f, err := os.Open("logs.xlsx")
if err != nil {
	return err
}
defer f.Close()
st, _ := f.Stat()

var count int
_, err = sheetio.ReadRows(f, st.Size(), nil, func(sheet string, row int, cells []sheetio.Cell) error {
	if sheet != "紀錄" {
		return sheetio.ErrSkipSheet // 略過這張表剩下的列
	}
	if row == 0 {
		return nil // 標題列
	}
	// callback 返回後 cells 會被重複使用，要保留的資料請自行複製
	level := ""
	if len(cells) > 1 {
		level = cells[1].String()
	}
	if level == "ERROR" {
		count++
	}
	if count > 1000 {
		return errStop // 回傳其他錯誤會停止讀取，並原樣回傳
	}
	return nil
})
if err != nil && !errors.Is(err, errStop) {
	return err
}
```

- `row` 是 0 起算的列號。空白列不會交付，所以列號可能不連續。
- 每列尾端的空白格已去除（每列的 `len(cells)` 可能不同）。
- 完全空白的工作表不會出現在 callback 裡；需要工作表清單時請用 `Read`。
- 因為記憶體用量固定，預設上限比 `Read` 寬（1 億格、單一項目 2 GB）。處理使用者上傳的檔案時，
  請自行設定 `Limits`。

## Gin 範例

### 下載（資料量中小）

```go
r.GET("/report", func(c *gin.Context) {
	f := sheetio.FormatXLSX
	if c.Query("format") == "ods" {
		f = sheetio.FormatODS
	}

	wb := sheetio.New()
	s := wb.AddSheet("報表")
	s.AddRow(sheetio.Str("ID"), sheetio.Str("姓名"))
	for _, u := range loadUsers() {
		s.AddRow(sheetio.Int(u.ID), sheetio.Str(u.Name))
	}

	var buf bytes.Buffer // 先組好，出錯時還能回 500
	if err := sheetio.Write(&buf, wb, f); err != nil {
		c.String(http.StatusInternalServerError, err.Error())
		return
	}
	name := "報表_" + time.Now().Format("20060102") + f.Ext()
	// 中文檔名用 RFC 5987 編碼
	c.Header("Content-Disposition", "attachment; filename*=UTF-8''"+url.PathEscape(name))
	c.Data(http.StatusOK, f.ContentType(), buf.Bytes())
})
```

### 下載（大量資料，串流）

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
			log.Println(err) // header 已送出，下載的檔案會不完整
			return
		}
		sw.WriteRow(sheetio.Date(t), sheetio.Str(level), sheetio.Str(msg))
	}
	if err := sw.Close(); err != nil {
		log.Println(err)
	}
})
```

### 上傳

```go
r.POST("/import", func(c *gin.Context) {
	c.Request.Body = http.MaxBytesReader(c.Writer, c.Request.Body, 20<<20)
	fh, err := c.FormFile("file")
	if err != nil {
		c.String(http.StatusBadRequest, "缺少檔案")
		return
	}
	f, err := fh.Open() // multipart.File 實作 io.ReaderAt，不需要先存檔
	if err != nil {
		c.String(http.StatusBadRequest, "無法開啟檔案")
		return
	}
	defer f.Close()

	opt := &sheetio.ReadOptions{Limits: sheetio.Limits{MaxFileSize: 20 << 20, MaxRows: 100000, MaxCols: 50}}
	var users []User
	_, err = sheetio.ReadRows(f, fh.Size, opt, func(sheet string, row int, cells []sheetio.Cell) error {
		if row == 0 {
			return nil // 標題列
		}
		if len(cells) < 2 {
			return fmt.Errorf("第 %d 列：需要 2 欄", row+1)
		}
		users = append(users, User{ID: int64(cells[0].Num), Name: cells[1].String()})
		return nil
	})
	switch {
	case errors.Is(err, sheetio.ErrUnsupported):
		c.String(http.StatusBadRequest, "只接受 .xlsx 與 .ods")
		return
	case errors.Is(err, sheetio.ErrLimit), errors.Is(err, sheetio.ErrCorrupt):
		c.String(http.StatusBadRequest, "檔案過大或已損毀")
		return
	case err != nil:
		c.String(http.StatusBadRequest, err.Error())
		return
	}
	c.JSON(http.StatusOK, gin.H{"imported": len(users)})
})
```

## 錯誤

| 錯誤 | 發生時機 |
|---|---|
| `ErrInvalid` | 寫出：活頁簿或某一列不合法（工作表名稱錯誤、NaN、日期超出範圍…） |
| `ErrClosed` | 寫出：`StreamWriter` 已經 `Close` 後又被使用 |
| `ErrUnsupported` | 讀取：不是 xlsx/ods（包含舊版 .xls 與加密檔） |
| `ErrCorrupt` | 讀取：檔案損毀或不符格式 |
| `ErrLimit` | 讀取：超過設定的上限 |
| `ErrSkipSheet` | 由 `ReadRows` 的 callback 回傳，略過該工作表剩下的列 |

錯誤都有包裝過，請用 `errors.Is` 判斷。

## 參考

### 儲存格

| 建構函式 | 說明 |
|---|---|
| `Str(s)` | 一律是文字格；開頭是 `=` 也不會被當成公式 |
| `Num(f)`、`Int(i)` | 以 float64 儲存。超過 2^53 的整數會失去精度：身分證、卡號這類資料請用 `Str` |
| `Bool(b)` | |
| `Date(t)` | 取牆上時間，不帶時區。範圍 1900-03-01 ～ 9999-12-31。自動套用日期或日期時間格式 |
| `Empty()` | 空白格，可搭配 `WithStyle` 做出有底色／框線的空格 |
| `.WithStyle(id)` | 套用 `AddStyle` 回傳的樣式 |

### 樣式

```go
sheetio.Style{
	Bold:      true,
	FillColor: "FFF2CC", // RRGGBB
	Border:    true,     // 四邊細框線
	Align:     sheetio.AlignCenter, // AlignLeft、AlignCenter、AlignRight
	Wrap:      true,
	NumFmt:    sheetio.NumFmtDecimal2,
}
```

數字格式：`NumFmtGeneral`、`NumFmtInteger`（`0`）、`NumFmtDecimal2`（`0.00`）、
`NumFmtThousands`（`#,##0`）、`NumFmtDate`（`yyyy-mm-dd`）、`NumFmtDateTime`（`yyyy-mm-dd hh:mm:ss`）。

### 工作表版面

| 方法／欄位 | 說明 |
|---|---|
| `SetColWidth(col, width)` | 欄寬，單位是字元數（0–255） |
| `SetRowHeight(row, height)` | 列高，單位是點（1–409）；只套用在有寫出的列 |
| `Merge(r1, c1, r2, c2)` | 合併範圍不可重疊 |
| `Freeze(rows, cols)` | 凍結上方幾列、左側幾欄 |
| `SetAutoFilter(headerRow, firstCol, lastCol)` | 在標題列加篩選按鈕，範圍自動延伸到最後一列資料 |

`StreamWriter` 用 `SheetDef` 設定同樣的內容（`ColWidths`、`RowHeights`、`Merges`、`FreezeRows`、
`FreezeCols`、`AutoFilter`）。

### 讀取預設上限

| 欄位 | `Read` | `ReadRows` |
|---|---|---|
| `MaxFileSize`（壓縮後） | 50 MB | 50 MB |
| `MaxEntries`（zip 項目數） | 10,000 | 10,000 |
| `MaxEntrySize`（單一項目解壓後） | 256 MB | 2 GB |
| `MaxTotalSize`（全部解壓後） | 512 MB | 4 GB |
| `MaxSheets` | 256 | 256 |
| `MaxRows` / `MaxCols` | 1,048,576 / 16,384 | 相同 |
| `MaxCells` | 5,000,000 | 100,000,000 |
| `MaxCellChars` | 32,767 | 相同 |

### 輔助函式

`ColName(2)` → `"C"`、`CellRef(0, 2)` → `"C1"`、`Detect(r, size)` 只判斷格式不解析、
`ValidateSheetName(name)`、`Format.Ext()` / `Format.ContentType()` / `Format.String()`。

## 效能

每列 10 欄（整數、字串、日期、小數、布林混合），在 2 核心的雲端容器上量測，峰值記憶體是整個行程的 RSS。

| 操作 | 10 萬列 | 100 萬列 | 100 萬列峰值記憶體 |
|---|---|---|---|
| 串流寫 xlsx | 1.0 秒 | 9.7 秒 | 8 MB |
| 串流讀 xlsx（`ReadRows`） | 2.4 秒 | 23.7 秒 | 11 MB |
| 串流寫 ods | 1.2 秒 | 12.0 秒 | 104 MB（壓縮後的 content.xml） |
| 串流讀 ods（`ReadRows`） | 4.6 秒 | 47 秒 | 9 MB |
| 整份讀 xlsx（`Read`） | 3.1 秒 | — | 151 MB（10 萬列） |

`Read` 會替每一格（包含列中間的空格）建立 `Cell`。大型或稀疏的表格用 `ReadRows` 比較快，
記憶體也少很多。

## 限制

- **公式**不會計算。讀取時拿到的是存檔程式留下的快取值；由不計算公式的工具產生的檔案，公式格會是空的。
- **數字**一律是 float64。身分證、卡號這類「看起來像數字」的資料請用 `Str` 寫出，才不會掉前導 0。
- **日期**不帶時區：寫出時取牆上時間，讀入時放在 `ReadOptions.Location`。
- **樣式、欄寬、列高、合併、凍結窗格、篩選只會寫出，不會讀回**。讀取只取值（合併範圍讀到左上角的值）。
- **不支援**：圖表、圖片、樞紐分析、條件格式、註解、資料驗證、舊版 `.xls`、加密檔。
- **ods 欄寬**是用 Excel 字元寬換算成公分，跟 xlsx 的效果會有些微差異。
- **ods 凍結窗格**寫在 `settings.xml`。Excel 開啟 ods 時不會套用凍結窗格（篩選與列高正常）。
- `StreamWriter` 一次寫一張工作表，不能多張交錯寫入。

## 授權

[MIT](LICENSE)
