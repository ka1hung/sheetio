package sheetio

import (
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"strconv"
	"strings"
	"time"
)

type xlsxRel struct{ typ, target string }

func readXLSX(a *zipArchive, rd *reader, h rowHandler) error {
	// 1. find the main workbook part
	pkgRels, err := readRels(a, "_rels/.rels", "")
	if err != nil {
		return err
	}
	wbPath := ""
	for _, r := range pkgRels {
		if strings.HasSuffix(r.typ, "/officeDocument") {
			wbPath = r.target
			break
		}
	}
	if wbPath == "" {
		wbPath = "xl/workbook.xml"
	}
	dir, base := path.Split(wbPath)
	rels, err := readRels(a, dir+"_rels/"+base+".rels", dir)
	if err != nil {
		return err
	}

	// 2. workbook.xml: list of sheets and the date system
	type sheetRef struct{ name, rid string }
	var refs []sheetRef
	date1904 := false
	err = scanXML(a, wbPath, func(d *xml.Decoder, se xml.StartElement) error {
		switch se.Name.Local {
		case "sheet":
			var s sheetRef
			for _, at := range se.Attr {
				switch {
				case at.Name.Local == "name":
					s.name = at.Value
				case at.Name.Local == "id" && strings.Contains(at.Name.Space, "relationships"):
					s.rid = at.Value
				}
			}
			if len(refs) >= rd.lim.MaxSheets {
				return fmt.Errorf("%w: more than %d sheets", ErrLimit, rd.lim.MaxSheets)
			}
			refs = append(refs, s)
		case "workbookPr":
			for _, at := range se.Attr {
				if at.Name.Local == "date1904" && (at.Value == "1" || at.Value == "true") {
					date1904 = true
				}
			}
		}
		return nil
	})
	if err != nil {
		return err
	}

	// 3. shared strings and styles
	var sst []string
	var dateXF []bool
	for _, r := range rels {
		switch {
		case strings.HasSuffix(r.typ, "/sharedStrings") && a.has(r.target):
			if sst, err = readSST(a, r.target, rd); err != nil {
				return err
			}
		case strings.HasSuffix(r.typ, "/styles") && a.has(r.target):
			if dateXF, err = readXLSXStyles(a, r.target); err != nil {
				return err
			}
		}
	}

	// 4. each sheet
	ctx := &xlsxCtx{rd: rd, sst: sst, dateXF: dateXF, date1904: date1904}
	n := 0
	for _, ref := range refs {
		r, ok := rels[ref.rid]
		if !ok || !strings.HasSuffix(r.typ, "/worksheet") {
			continue // skip chartsheets, dialogsheets, etc.
		}
		if err := rd.startSheet(h, ref.name); err != nil {
			return err
		}
		if err := ctx.readSheet(a, r.target, ref.name, h); err != nil {
			return err
		}
		if err := h.endSheet(); err != nil {
			return err
		}
		n++
	}
	if n == 0 {
		return fmt.Errorf("%w: no worksheets", ErrCorrupt)
	}
	return nil
}

// scanXML visits each StartElement.
func scanXML(a *zipArchive, name string, fn func(*xml.Decoder, xml.StartElement) error) error {
	rc, err := a.open(name)
	if err != nil {
		return err
	}
	defer rc.Close()
	d := xml.NewDecoder(rc)
	for {
		tok, err := d.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return wrapXMLErr(name, err)
		}
		if se, ok := tok.(xml.StartElement); ok {
			if err := fn(d, se); err != nil {
				return err
			}
		}
	}
}

func wrapXMLErr(name string, err error) error {
	if isLimitErr(err) {
		return err
	}
	return fmt.Errorf("%w: %s: %v", ErrCorrupt, name, err)
}

func isLimitErr(err error) bool {
	for e := err; e != nil; {
		if e == ErrLimit {
			return true
		}
		u, ok := e.(interface{ Unwrap() error })
		if !ok {
			return false
		}
		e = u.Unwrap()
	}
	return false
}

func readRels(a *zipArchive, name, baseDir string) (map[string]xlsxRel, error) {
	out := map[string]xlsxRel{}
	if !a.has(name) {
		return out, nil
	}
	err := scanXML(a, name, func(_ *xml.Decoder, se xml.StartElement) error {
		if se.Name.Local != "Relationship" {
			return nil
		}
		var id, typ, target, mode string
		for _, at := range se.Attr {
			switch at.Name.Local {
			case "Id":
				id = at.Value
			case "Type":
				typ = at.Value
			case "Target":
				target = at.Value
			case "TargetMode":
				mode = at.Value
			}
		}
		if mode == "External" || len(out) > 100000 {
			return nil
		}
		out[id] = xlsxRel{typ: typ, target: resolvePart(baseDir, target)}
		return nil
	})
	return out, err
}

// resolvePart turns a relationship target into a path inside the zip (an in-memory lookup only, no file system access).
func resolvePart(baseDir, target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(path.Clean(target), "/")
	}
	p := path.Clean(path.Join(baseDir, target))
	return strings.TrimPrefix(p, "/")
}

func readSST(a *zipArchive, name string, rd *reader) ([]string, error) {
	rc, err := a.open(name)
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	d := xml.NewDecoder(rc)
	var out []string
	var buf textBuf
	inSI, inT, rph := false, false, 0
	for {
		tok, err := d.RawToken()
		if err == io.EOF {
			return out, nil
		}
		if err != nil {
			return nil, wrapXMLErr(name, err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "si":
				inSI = true
				buf.reset(rd.lim.MaxCellChars)
			case "t":
				inT = inSI && rph == 0
			case "rPh": // phonetic text (furigana), not part of the cell value
				rph++
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "si":
				if len(out) >= rd.lim.MaxCells {
					return nil, fmt.Errorf("%w: too many shared strings", ErrLimit)
				}
				out = append(out, buf.String())
				inSI = false
			case "t":
				inT = false
			case "rPh":
				rph--
			}
		case xml.CharData:
			if inT {
				if err := buf.write(string(t)); err != nil {
					return nil, err
				}
			}
		}
	}
}

// readXLSXStyles reports, for each cellXfs index, whether it is a date format.
func readXLSXStyles(a *zipArchive, name string) ([]bool, error) {
	custom := map[int]string{}
	var xfFmt []int
	inXfs := false
	rc, err := a.open(name)
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	d := xml.NewDecoder(rc)
	for {
		tok, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, wrapXMLErr(name, err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "numFmt":
				var id int
				var code string
				for _, at := range t.Attr {
					switch at.Name.Local {
					case "numFmtId":
						id, _ = strconv.Atoi(at.Value)
					case "formatCode":
						code = at.Value
					}
				}
				if len(custom) < 10000 {
					custom[id] = code
				}
			case "cellXfs":
				inXfs = true
			case "xf":
				if inXfs {
					if len(xfFmt) >= 100000 {
						return nil, fmt.Errorf("%w: too many cell formats", ErrLimit)
					}
					id := 0
					for _, at := range t.Attr {
						if at.Name.Local == "numFmtId" {
							id, _ = strconv.Atoi(at.Value)
						}
					}
					xfFmt = append(xfFmt, id)
				}
			}
		case xml.EndElement:
			if t.Name.Local == "cellXfs" {
				inXfs = false
			}
		}
	}
	out := make([]bool, len(xfFmt))
	for i, id := range xfFmt {
		if code, ok := custom[id]; ok {
			out[i] = isDateFormatCode(code)
		} else {
			out[i] = isBuiltinDateFmt(id)
		}
	}
	return out, nil
}

func isBuiltinDateFmt(id int) bool {
	switch {
	case id >= 14 && id <= 22, id >= 45 && id <= 47:
		return true
	case id >= 27 && id <= 36, id >= 50 && id <= 58: // East Asian date formats (including the zh-TW Minguo calendar)
		return true
	}
	return false
}

// isDateFormatCode reports whether a custom format code is a date or time format.
func isDateFormatCode(code string) bool {
	if i := strings.IndexByte(code, ';'); i >= 0 {
		code = code[:i]
	}
	inQuote, inBracket := false, false
	for i := 0; i < len(code); i++ {
		c := code[i]
		switch {
		case inQuote:
			if c == '"' {
				inQuote = false
			}
		case inBracket:
			if c == ']' {
				inBracket = false
			}
		case c == '"':
			inQuote = true
		case c == '[':
			// [h] [mm] [ss] are elapsed time, treated as time; other brackets such as [$-404] and [Red] are skipped
			rest := strings.ToLower(code[i:])
			if strings.HasPrefix(rest, "[h") || strings.HasPrefix(rest, "[m") || strings.HasPrefix(rest, "[s") {
				if !strings.HasPrefix(rest, "[magenta") {
					return true
				}
			}
			inBracket = true
		case c == '\\' || c == '_' || c == '*':
			i++ // skip the escaped or fill character
		default:
			switch c | 0x20 { // lowercase
			case 'y', 'm', 'd', 'h', 's', 'e', 'g':
				if c|0x20 == 'g' && strings.HasPrefix(strings.ToLower(code[i:]), "general") {
					return false
				}
				if c|0x20 == 'e' && i > 0 && (code[i-1] == '0' || code[i-1] == '#' || code[i-1] == '.') {
					continue // scientific notation 0.00E+00
				}
				return true
			}
		}
	}
	return false
}

type xlsxCtx struct {
	rd       *reader
	sst      []string
	dateXF   []bool
	date1904 bool
}

func (x *xlsxCtx) readSheet(a *zipArchive, name, sheet string, h rowHandler) error {
	rc, err := a.open(name)
	if err != nil {
		return err
	}
	defer rc.Close()
	d := xml.NewDecoder(rc)
	lim := x.rd.lim

	row, col := -1, -1
	var (
		inCell, inV, inT, inIs bool
		hasV                   bool
		cType                  string
		cStyle                 int
		cCol                   int
		buf                    textBuf
		rowBuf                 []Cell
	)
	emit := func() error {
		if row < 0 || len(rowBuf) == 0 {
			return nil
		}
		err := x.rd.emit(h, sheet, row, rowBuf)
		rowBuf = rowBuf[:0]
		return err
	}
	for {
		tok, err := d.RawToken()
		if err == io.EOF {
			err = emit()
			if err == ErrSkipSheet {
				return nil
			}
			return err
		}
		if err != nil {
			return wrapXMLErr(name, err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "row":
				if err := emit(); err != nil {
					if err == ErrSkipSheet {
						return nil
					}
					return err
				}
				next := row + 1
				for _, at := range t.Attr {
					if at.Name.Local == "r" {
						n, err := strconv.Atoi(at.Value)
						if err != nil || n < 1 {
							return fmt.Errorf("%w: %s: bad row number %q", ErrCorrupt, name, at.Value)
						}
						if n > lim.MaxRows {
							return fmt.Errorf("%w: sheet %q row %d exceeds %d", ErrLimit, sheet, n, lim.MaxRows)
						}
						next = n - 1
					}
				}
				if next >= lim.MaxRows {
					return fmt.Errorf("%w: sheet %q exceeds %d rows", ErrLimit, sheet, lim.MaxRows)
				}
				row, col = next, -1
			case "c":
				if row < 0 {
					row = 0
				}
				inCell, hasV = true, false
				cType, cStyle = "n", 0
				cCol = col + 1
				buf.reset(lim.MaxCellChars)
				for _, at := range t.Attr {
					switch at.Name.Local {
					case "r":
						r, c, ok := parseCellRef(at.Value)
						if !ok {
							return fmt.Errorf("%w: %s: bad cell reference %q", ErrCorrupt, name, at.Value)
						}
						if r >= lim.MaxRows {
							return fmt.Errorf("%w: sheet %q cell %s out of range", ErrLimit, sheet, at.Value)
						}
						cCol = c
					case "t":
						cType = at.Value
					case "s":
						cStyle, _ = strconv.Atoi(at.Value)
					}
				}
				if cCol >= lim.MaxCols {
					return fmt.Errorf("%w: sheet %q column %d exceeds %d", ErrLimit, sheet, cCol+1, lim.MaxCols)
				}
				col = cCol
			case "v":
				inV, hasV = inCell, inCell
			case "is":
				inIs, hasV = inCell, inCell
			case "t":
				inT = inIs
			case "rPh", "f", "extLst":
				// phonetics, formula text and extensions are not needed; only the cached value <v> is used
				if err := rawSkip(d); err != nil {
					return wrapXMLErr(name, err)
				}
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "v":
				inV = false
			case "t":
				inT = false
			case "is":
				inIs = false
			case "row":
				if err := emit(); err != nil {
					if err == ErrSkipSheet {
						return nil
					}
					return err
				}
			case "c":
				inCell = false
				if !hasV {
					continue
				}
				cell, err := x.convert(cType, buf.String(), cStyle)
				if err != nil {
					return fmt.Errorf("sheet %q cell %s: %w", sheet, CellRef(row, cCol), err)
				}
				if cell.IsEmpty() {
					continue
				}
				for len(rowBuf) <= cCol {
					rowBuf = append(rowBuf, Cell{})
				}
				rowBuf[cCol] = cell
			}
		case xml.CharData:
			if inV || inT {
				if err := buf.write(string(t)); err != nil {
					return err
				}
			}
		}
	}
}

func (x *xlsxCtx) convert(typ, v string, style int) (Cell, error) {
	switch typ {
	case "s":
		i, err := strconv.Atoi(strings.TrimSpace(v))
		if err != nil || i < 0 || i >= len(x.sst) {
			return Cell{}, fmt.Errorf("%w: shared string index %q out of range", ErrCorrupt, v)
		}
		return Str(x.sst[i]), nil
	case "inlineStr", "str", "e":
		if v == "" {
			return Cell{}, nil
		}
		return Str(v), nil
	case "b":
		return Bool(strings.TrimSpace(v) == "1" || strings.EqualFold(strings.TrimSpace(v), "true")), nil
	case "d":
		if t, ok := parseISODate(strings.TrimSpace(v), x.rd.loc); ok {
			return Date(t), nil
		}
		return Str(v), nil
	default: // "n"
		v = strings.TrimSpace(v)
		if v == "" {
			return Cell{}, nil
		}
		f, err := strconv.ParseFloat(v, 64)
		if err != nil {
			return Cell{}, fmt.Errorf("%w: bad number %q", ErrCorrupt, v)
		}
		if style >= 0 && style < len(x.dateXF) && x.dateXF[style] {
			if t, ok := fromExcelSerial(f, x.date1904, x.rd.loc); ok {
				return Date(t), nil
			}
		}
		return Num(f), nil
	}
}

var isoLayouts = []string{
	"2006-01-02T15:04:05.999999999",
	"2006-01-02T15:04:05",
	"2006-01-02T15:04",
	"2006-01-02",
}

func parseISODate(s string, loc *time.Location) (time.Time, bool) {
	if len(s) > 64 {
		return time.Time{}, false
	}
	// values with a time zone (such as Z or +08:00) are converted to wall-clock time first
	if t, err := time.Parse(time.RFC3339Nano, s); err == nil {
		w := wallUTC(t)
		return time.Date(w.Year(), w.Month(), w.Day(), w.Hour(), w.Minute(), w.Second(), w.Nanosecond(), loc), true
	}
	for _, l := range isoLayouts {
		if t, err := time.ParseInLocation(l, s, loc); err == nil {
			return t, true
		}
	}
	return time.Time{}, false
}
