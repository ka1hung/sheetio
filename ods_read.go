package sheetio

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"
)

const (
	nsOffice = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
	nsTable  = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
	nsText   = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
)

type odsPending struct {
	col  int
	cell Cell
}

func readODS(a *zipArchive, rd *reader, h rowHandler) error {
	const name = "content.xml"
	rc, err := a.open(name)
	if err != nil {
		return err
	}
	defer rc.Close()
	d := xml.NewDecoder(rc)
	lim := rd.lim
	nsheets := 0
	ns := nsResolver{}

	var (
		inSheet   bool
		sheetName string
		skipping  bool
		rowBuf    []Cell
		row, col  int
		rowRepeat int
		pending   []odsPending

		inCell     bool
		colRepeat  int
		valueType  string
		value      string
		dateValue  string
		boolValue  string
		strValue   string
		hasStrAttr bool
		buf        textBuf
		paras      int
		inP        bool
		lastSpace  bool // for ODF whitespace collapsing
	)

	attr := func(se xml.StartElement, space, local string) (string, bool) {
		for _, at := range se.Attr {
			if at.Name.Local == local && at.Name.Space == space {
				return at.Value, true
			}
		}
		return "", false
	}
	repeat := func(se xml.StartElement, local string, max int) (int, error) {
		v, ok := attr(se, nsTable, local)
		if !ok {
			return 1, nil
		}
		n, err := strconv.Atoi(v)
		if err != nil || n < 1 {
			return 0, fmt.Errorf("%w: %s: bad %s %q", ErrCorrupt, name, local, v)
		}
		if n > max+1 {
			n = max + 1 // we only need to know it is over the limit; this also avoids integer overflow later
		}
		return n, nil
	}

	for {
		tok, err := d.RawToken()
		if err == io.EOF {
			break
		}
		if err != nil {
			return wrapXMLErr(name, err)
		}
		switch t := tok.(type) {
		case xml.StartElement:
			ns.start(&t)
			tok = t
		case xml.EndElement:
			ns.end(&t)
			tok = t
		}
		if skipping { // skip the rest of the current sheet
			if et, ok := tok.(xml.EndElement); ok && et.Name.Space == nsTable && et.Name.Local == "table" {
				skipping, inSheet = false, false
				if err := h.endSheet(); err != nil {
					return err
				}
			}
			continue
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Space {
			case nsTable:
				switch t.Name.Local {
				case "table":
					if inCell || inSheet { // nested sub-table, skip
						if err := rawSkip(d); err != nil {
							return wrapXMLErr(name, err)
						}
						continue
					}
					sheetName, _ = attr(t, nsTable, "name")
					if err := rd.startSheet(h, sheetName); err != nil {
						return err
					}
					inSheet, row = true, 0
					nsheets++
				case "table-row":
					if !inSheet {
						continue
					}
					if rowRepeat, err = repeat(t, "number-rows-repeated", MaxSheetRows); err != nil {
						return err
					}
					col, pending = 0, pending[:0]
				case "table-cell", "covered-table-cell":
					if !inSheet {
						continue
					}
					if colRepeat, err = repeat(t, "number-columns-repeated", MaxSheetCols); err != nil {
						return err
					}
					inCell, paras, inP = true, 0, false
					valueType, _ = attr(t, nsOffice, "value-type")
					value, _ = attr(t, nsOffice, "value")
					dateValue, _ = attr(t, nsOffice, "date-value")
					boolValue, _ = attr(t, nsOffice, "boolean-value")
					strValue, hasStrAttr = attr(t, nsOffice, "string-value")
					buf.reset(lim.MaxCellChars)
				}
			case nsText:
				if !inCell {
					continue
				}
				switch t.Name.Local {
				case "p", "h":
					if paras > 0 {
						if err := buf.write("\n"); err != nil {
							return err
						}
					}
					paras++
					inP, lastSpace = true, true
				case "s":
					n := 1
					if v, ok := attr(t, nsText, "c"); ok {
						if n, err = strconv.Atoi(v); err != nil || n < 1 {
							n = 1
						}
					}
					if n > lim.MaxCellChars {
						return fmt.Errorf("%w: text:s count %d", ErrLimit, n)
					}
					if err := buf.write(strings.Repeat(" ", n)); err != nil {
						return err
					}
					lastSpace = false
				case "tab":
					if err := buf.write("\t"); err != nil {
						return err
					}
				case "line-break":
					if err := buf.write("\n"); err != nil {
						return err
					}
				case "note", "ruby-text", "bookmark-ref", "sequence-ref":
					if err := rawSkip(d); err != nil {
						return wrapXMLErr(name, err)
					}
				}
			case nsOffice:
				if t.Name.Local == "annotation" { // comments are not cell content
					if err := rawSkip(d); err != nil {
						return wrapXMLErr(name, err)
					}
				}
			default:
				// drawing objects such as draw:frame
				if inCell && t.Name.Space != "" && strings.Contains(t.Name.Space, "drawing") {
					if err := rawSkip(d); err != nil {
						return wrapXMLErr(name, err)
					}
				}
			}

		case xml.EndElement:
			switch {
			case t.Name.Space == nsText && (t.Name.Local == "p" || t.Name.Local == "h"):
				inP = false
			case t.Name.Space == nsTable && (t.Name.Local == "table-cell" || t.Name.Local == "covered-table-cell"):
				if !inCell {
					continue
				}
				inCell = false
				text := buf.String()
				if hasStrAttr {
					text = strValue
				}
				cell := odsCell(valueType, value, dateValue, boolValue, text, rd)
				if !cell.IsEmpty() {
					if col+colRepeat > lim.MaxCols {
						return fmt.Errorf("%w: sheet %q column exceeds %d", ErrLimit, sheetName, lim.MaxCols)
					}
					for i := 0; i < colRepeat; i++ {
						pending = append(pending, odsPending{col + i, cell})
						if len(pending) > lim.MaxCells {
							return fmt.Errorf("%w: more than %d cells", ErrLimit, lim.MaxCells)
						}
					}
				}
				col += colRepeat
				if col > MaxSheetCols*2 {
					col = MaxSheetCols * 2 // only blank cells can get here; clamp to avoid overflow
				}
			case t.Name.Space == nsTable && t.Name.Local == "table-row":
				if !inSheet {
					continue
				}
				if len(pending) > 0 {
					if row+rowRepeat > lim.MaxRows {
						return fmt.Errorf("%w: sheet %q row exceeds %d", ErrLimit, sheetName, lim.MaxRows)
					}
					rowBuf = rowBuf[:0]
					for _, p := range pending {
						for len(rowBuf) < p.col {
							rowBuf = append(rowBuf, Cell{})
						}
						rowBuf = append(rowBuf, p.cell)
					}
					for i := 0; i < rowRepeat && !skipping; i++ {
						if err := rd.emit(h, sheetName, row+i, rowBuf); err == ErrSkipSheet {
							skipping = true
						} else if err != nil {
							return err
						}
					}
				}
				row += rowRepeat
				if row > MaxSheetRows*2 {
					row = MaxSheetRows * 2
				}
			case t.Name.Space == nsTable && t.Name.Local == "table":
				if inSheet {
					inSheet = false
					if err := h.endSheet(); err != nil {
						return err
					}
				}
			}

		case xml.CharData:
			if !inCell || !inP {
				continue
			}
			// ODF whitespace rule: a run of whitespace collapses to one space, and leading whitespace in a paragraph is ignored
			var sb strings.Builder
			for _, r := range string(t) {
				if r == ' ' || r == '\t' || r == '\n' || r == '\r' {
					if !lastSpace {
						sb.WriteByte(' ')
						lastSpace = true
					}
					continue
				}
				sb.WriteRune(r)
				lastSpace = false
			}
			if err := buf.write(sb.String()); err != nil {
				return err
			}
		}
	}
	if nsheets == 0 {
		return fmt.Errorf("%w: no sheets in ods", ErrCorrupt)
	}
	return nil
}

func odsCell(valueType, value, dateValue, boolValue, text string, rd *reader) Cell {
	switch valueType {
	case "float", "percentage", "currency":
		if f, err := strconv.ParseFloat(strings.TrimSpace(value), 64); err == nil {
			return Num(f)
		}
	case "boolean":
		return Bool(boolValue == "true" || boolValue == "1")
	case "date":
		if t, ok := parseISODate(strings.TrimSpace(dateValue), rd.loc); ok {
			return Date(t)
		}
	}
	// string, time (PT12H30M00S) and any value that cannot be parsed use the displayed text
	if text == "" {
		return Cell{}
	}
	return Str(text)
}
