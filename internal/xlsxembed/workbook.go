package xlsxembed

import (
	"archive/zip"
	"bytes"
	"compress/flate"
	"encoding/xml"
	"fmt"
	"hash/crc32"
	"io"
	"path"
	"sort"
	"strconv"
	"strings"

	"why-pptx/internal/rels"
	"why-pptx/internal/xlref"
)

type CellValue struct {
	Number *float64
	String *string
}

type MissingNumericPolicy int

const (
	MissingNumericEmpty MissingNumericPolicy = iota
	MissingNumericZero
)

type Workbook struct {
	data    []byte
	reader  *zip.Reader
	index   map[string]*zip.File
	overlay map[string][]byte
	sheets  map[string]string
}

func Open(data []byte) (*Workbook, error) {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("open xlsx: %w", err)
	}

	index := make(map[string]*zip.File, len(reader.File))
	for _, part := range reader.File {
		index[part.Name] = part
	}

	wb := &Workbook{
		data:    data,
		reader:  reader,
		index:   index,
		overlay: make(map[string][]byte),
	}

	sheets, err := wb.loadSheets()
	if err != nil {
		return nil, err
	}
	wb.sheets = sheets

	return wb, nil
}

func (wb *Workbook) SetCell(sheetName, cellRef string, v CellValue) error {
	if wb == nil || wb.reader == nil {
		return fmt.Errorf("workbook not initialized")
	}
	if sheetName == "" {
		return fmt.Errorf("sheet name is required")
	}
	// Sheet names are matched exactly as stored in workbook.xml (Unicode supported).
	if (v.Number == nil && v.String == nil) || (v.Number != nil && v.String != nil) {
		return fmt.Errorf("cell value must specify exactly one of number or string")
	}

	col, row, normalized, err := xlref.SplitCellRef(cellRef)
	if err != nil {
		return err
	}

	sheetPath, ok := wb.sheets[sheetName]
	if !ok {
		return fmt.Errorf("sheet %q not found", sheetName)
	}

	data, err := wb.readPart(sheetPath)
	if err != nil {
		return fmt.Errorf("read sheet %q: %w", sheetPath, err)
	}

	update := cellUpdate{
		Ref:   normalized,
		Row:   row,
		Col:   col,
		Value: v,
	}

	updated, err := updateSheetXML(data, []cellUpdate{update})
	if err != nil {
		return fmt.Errorf("update sheet %q: %w", sheetPath, err)
	}

	wb.overlay[sheetPath] = updated
	return nil
}

func (wb *Workbook) Save() ([]byte, error) {
	if wb == nil || wb.reader == nil {
		return nil, fmt.Errorf("workbook not initialized")
	}

	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	written := make(map[string]struct{}, len(wb.reader.File)+len(wb.overlay))

	for _, part := range wb.reader.File {
		name := part.Name
		if data, ok := wb.overlay[name]; ok {
			if err := writeOverrideEntry(writer, part, data); err != nil {
				_ = writer.Close()
				return nil, fmt.Errorf("write part %q: %w", name, err)
			}
		} else {
			if err := writer.Copy(part); err != nil {
				_ = writer.Close()
				return nil, fmt.Errorf("copy part %q: %w", name, err)
			}
		}
		written[name] = struct{}{}
	}

	for name, data := range wb.overlay {
		if _, ok := written[name]; ok {
			continue
		}
		if err := writeNewEntry(writer, name, data); err != nil {
			_ = writer.Close()
			return nil, fmt.Errorf("write part %q: %w", name, err)
		}
	}

	if err := writer.Close(); err != nil {
		return nil, fmt.Errorf("close xlsx: %w", err)
	}

	return buf.Bytes(), nil
}

func (wb *Workbook) GetRangeValues(sheetName, startCell, endCell string, policy MissingNumericPolicy) ([]string, error) {
	if wb == nil || wb.reader == nil {
		return nil, fmt.Errorf("workbook not initialized")
	}
	if sheetName == "" {
		return nil, fmt.Errorf("sheet name is required")
	}

	sheetPath, ok := wb.sheets[sheetName]
	if !ok {
		return nil, fmt.Errorf("sheet %q not found", sheetName)
	}

	startCol, startRow, startRef, err := xlref.SplitCellRef(startCell)
	if err != nil {
		return nil, fmt.Errorf("invalid start cell %q: %w", startCell, err)
	}
	endCol, endRow, endRef, err := xlref.SplitCellRef(endCell)
	if err != nil {
		return nil, fmt.Errorf("invalid end cell %q: %w", endCell, err)
	}

	if startCol != endCol && startRow != endRow {
		return nil, fmt.Errorf("2D range %s:%s not supported", startRef, endRef)
	}

	if startCol == endCol && startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startRow == endRow && colToIndex(startCol) > colToIndex(endCol) {
		startCol, endCol = endCol, startCol
	}

	targets := make(map[string]struct{})
	var ordered []string
	if startCol == endCol {
		for row := startRow; row <= endRow; row++ {
			ref := fmt.Sprintf("%s%d", startCol, row)
			ordered = append(ordered, ref)
			targets[ref] = struct{}{}
		}
	} else {
		for col := colToIndex(startCol); col <= colToIndex(endCol); col++ {
			ref := fmt.Sprintf("%s%d", indexToCol(col), startRow)
			ordered = append(ordered, ref)
			targets[ref] = struct{}{}
		}
	}

	data, err := wb.readPart(sheetPath)
	if err != nil {
		return nil, fmt.Errorf("read sheet %q: %w", sheetPath, err)
	}

	values, err := readCellValues(data, targets, policy)
	if err != nil {
		return nil, err
	}

	out := make([]string, len(ordered))
	for i, ref := range ordered {
		out[i] = values[ref]
	}
	return out, nil
}

func (wb *Workbook) loadSheets() (map[string]string, error) {
	workbookData, err := wb.readPart("xl/workbook.xml")
	if err != nil {
		return nil, fmt.Errorf("read workbook.xml: %w", err)
	}
	relsData, err := wb.readPart("xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, fmt.Errorf("read workbook rels: %w", err)
	}

	sheets, err := parseWorkbookSheets(workbookData)
	if err != nil {
		return nil, err
	}

	parsedRels, err := rels.Parse(bytes.NewReader(relsData))
	if err != nil {
		return nil, err
	}

	sheetPaths := make(map[string]string, len(sheets))
	for name, relID := range sheets {
		rel, ok := parsedRels.Resolve(relID)
		if !ok {
			return nil, fmt.Errorf("sheet %q missing rel %q", name, relID)
		}
		target := rels.ResolveTarget("xl/workbook.xml", rel.Target)
		target = path.Clean(target)
		if !strings.HasPrefix(target, "xl/") {
			target = path.Join("xl", strings.TrimLeft(target, "/"))
		}
		sheetPaths[name] = target
	}

	return sheetPaths, nil
}

func (wb *Workbook) readPart(name string) ([]byte, error) {
	if data, ok := wb.overlay[name]; ok {
		return append([]byte(nil), data...), nil
	}
	part, ok := wb.index[name]
	if !ok {
		return nil, fmt.Errorf("part %q not found", name)
	}
	reader, err := part.Open()
	if err != nil {
		return nil, err
	}
	defer reader.Close()
	data, err := io.ReadAll(reader)
	if err != nil {
		return nil, err
	}
	return data, nil
}

func parseWorkbookSheets(data []byte) (map[string]string, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	sheets := make(map[string]string)

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse workbook: %w", err)
		}

		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "sheet" {
			continue
		}

		var name string
		var relID string
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "name":
				name = attr.Value
			case "id":
				if attr.Name.Space != "" {
					relID = attr.Value
				}
			}
		}
		if name != "" && relID != "" {
			sheets[name] = relID
		}
	}

	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in workbook")
	}

	return sheets, nil
}

type cellUpdate struct {
	Ref   string
	Row   int
	Col   string
	Value CellValue
}

func updateSheetXML(data []byte, updates []cellUpdate) ([]byte, error) {
	if len(updates) == 0 {
		return data, nil
	}

	updateByRef := make(map[string]cellUpdate, len(updates))
	updatesByRow := make(map[int][]cellUpdate)
	for _, update := range updates {
		updateByRef[update.Ref] = update
		updatesByRow[update.Row] = append(updatesByRow[update.Row], update)
	}

	pending := make(map[string]cellUpdate, len(updateByRef))
	for ref, update := range updateByRef {
		pending[ref] = update
	}

	seenRows := make(map[int]bool)
	var rowName xml.Name
	var cellName xml.Name
	var currentRow int
	rowPending := map[string]cellUpdate(nil)
	foundSheetData := false

	decoder := xml.NewDecoder(bytes.NewReader(data))
	var buf bytes.Buffer
	encoder := xml.NewEncoder(&buf)

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse worksheet: %w", err)
		}

		switch tok := token.(type) {
		case xml.StartElement:
			if tok.Name.Local == "sheetData" {
				foundSheetData = true
			}
			if tok.Name.Local == "row" {
				if rowName.Local == "" {
					rowName = tok.Name
				}
				currentRow = parseRowNumber(tok.Attr)
				if currentRow > 0 {
					seenRows[currentRow] = true
					rowPending = make(map[string]cellUpdate)
					for _, update := range updatesByRow[currentRow] {
						if _, ok := pending[update.Ref]; ok {
							rowPending[update.Ref] = update
						}
					}
				}
				if err := encoder.EncodeToken(tok); err != nil {
					return nil, err
				}
				continue
			}

			if tok.Name.Local == "c" && currentRow > 0 {
				if cellName.Local == "" {
					cellName = tok.Name
				}
				cellRef := cellRefFromAttrs(tok.Attr)
				if cellRef != "" {
					normalized, err := xlref.NormalizeCellRef(cellRef)
					if err == nil {
						if update, ok := pending[normalized]; ok {
							delete(pending, normalized)
							if rowPending != nil {
								delete(rowPending, normalized)
							}
							if err := writeUpdatedCell(decoder, encoder, tok, normalized, update.Value); err != nil {
								return nil, err
							}
							continue
						}
					}
				}
			}

			if err := encoder.EncodeToken(tok); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if tok.Name.Local == "row" {
				if len(rowPending) > 0 {
					writePendingCells(encoder, cellName, rowPending)
					for ref := range rowPending {
						delete(pending, ref)
					}
				}
				if err := encoder.EncodeToken(tok); err != nil {
					return nil, err
				}
				currentRow = 0
				rowPending = nil
				continue
			}

			if tok.Name.Local == "sheetData" {
				if len(pending) > 0 {
					appendMissingRows(encoder, rowName, cellName, pending, seenRows)
				}
				if err := encoder.EncodeToken(tok); err != nil {
					return nil, err
				}
				continue
			}

			if err := encoder.EncodeToken(tok); err != nil {
				return nil, err
			}
		default:
			if err := encoder.EncodeToken(tok); err != nil {
				return nil, err
			}
		}
	}

	if err := encoder.Flush(); err != nil {
		return nil, err
	}
	if !foundSheetData {
		return nil, fmt.Errorf("worksheet missing sheetData")
	}

	return buf.Bytes(), nil
}

func parseRowNumber(attrs []xml.Attr) int {
	for _, attr := range attrs {
		if attr.Name.Local == "r" {
			row, err := strconv.Atoi(attr.Value)
			if err == nil && row > 0 {
				return row
			}
		}
	}
	return 0
}

func cellRefFromAttrs(attrs []xml.Attr) string {
	for _, attr := range attrs {
		if attr.Name.Local == "r" {
			return attr.Value
		}
	}
	return ""
}

func skipElement(decoder *xml.Decoder) error {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return err
		}
		switch token.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}

func writeUpdatedCell(decoder *xml.Decoder, encoder *xml.Encoder, start xml.StartElement, cellRef string, value CellValue) error {
	start.Attr = buildCellAttrs(cellRef, start.Attr, value)
	if err := encoder.EncodeToken(start); err != nil {
		return err
	}

	depth := 1
	wroteValue := false

	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return err
		}

		switch tok := token.(type) {
		case xml.StartElement:
			if depth == 1 && (tok.Name.Local == "v" || tok.Name.Local == "is") {
				if wroteValue {
					if err := skipElement(decoder); err != nil {
						return err
					}
					continue
				}
				if tok.Name.Local == "v" && value.Number != nil {
					if err := writeNumberValue(encoder, *value.Number); err != nil {
						return err
					}
					wroteValue = true
					if err := skipElement(decoder); err != nil {
						return err
					}
					continue
				}
				if tok.Name.Local == "is" && value.String != nil {
					if err := writeInlineStr(encoder, *value.String); err != nil {
						return err
					}
					wroteValue = true
					if err := skipElement(decoder); err != nil {
						return err
					}
					continue
				}
				if err := skipElement(decoder); err != nil {
					return err
				}
				continue
			}
			if err := encoder.EncodeToken(tok); err != nil {
				return err
			}
			depth++
		case xml.EndElement:
			depth--
			if depth == 0 {
				if !wroteValue {
					if value.String != nil {
						if err := writeInlineStr(encoder, *value.String); err != nil {
							return err
						}
					} else if value.Number != nil {
						if err := writeNumberValue(encoder, *value.Number); err != nil {
							return err
						}
					}
				}
				if err := encoder.EncodeToken(tok); err != nil {
					return err
				}
				return nil
			}
			if err := encoder.EncodeToken(tok); err != nil {
				return err
			}
		default:
			if err := encoder.EncodeToken(tok); err != nil {
				return err
			}
		}
	}

	return nil
}

func readCellValues(data []byte, targets map[string]struct{}, policy MissingNumericPolicy) (map[string]string, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	values := make(map[string]string, len(targets))

	var inCell bool
	var cellRef string
	var cellType string
	var inValue bool
	var valueBuf strings.Builder

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse worksheet: %w", err)
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "c":
				cellRef = ""
				cellType = ""
				inCell = false
				for _, attr := range tok.Attr {
					if attr.Name.Local == "r" {
						cellRef = attr.Value
					} else if attr.Name.Local == "t" {
						cellType = attr.Value
					}
				}
				if cellRef != "" {
					normalized, err := xlref.NormalizeCellRef(cellRef)
					if err == nil {
						if _, ok := targets[normalized]; ok {
							if cellType != "" && cellType != "n" && cellType != "inlineStr" {
								return nil, fmt.Errorf("unsupported cell type %q at %s", cellType, normalized)
							}
							cellRef = normalized
							inCell = true
						}
					}
				}
			case "v":
				if inCell && (cellType == "" || cellType == "n") {
					inValue = true
					valueBuf.Reset()
				}
			case "t":
				if inCell && cellType == "inlineStr" {
					inValue = true
					valueBuf.Reset()
				}
			}
		case xml.EndElement:
			switch tok.Name.Local {
			case "c":
				if inCell {
					val := valueBuf.String()
					if cellType == "" || cellType == "n" {
						val = strings.TrimSpace(val)
					}
					values[cellRef] = val
				}
				inCell = false
				inValue = false
				cellRef = ""
				cellType = ""
			case "v", "t":
				inValue = false
			}
		case xml.CharData:
			if inCell && inValue {
				valueBuf.Write([]byte(tok))
			}
		}
	}

	for ref := range targets {
		if _, ok := values[ref]; !ok {
			if policy == MissingNumericZero {
				values[ref] = "0"
			} else {
				values[ref] = ""
			}
		}
	}

	return values, nil
}

func writePendingCells(encoder *xml.Encoder, cellName xml.Name, pending map[string]cellUpdate) {
	if len(pending) == 0 {
		return
	}
	if cellName.Local == "" {
		cellName = xml.Name{Local: "c"}
	}

	updates := make([]cellUpdate, 0, len(pending))
	for _, update := range pending {
		updates = append(updates, update)
	}
	sort.Slice(updates, func(i, j int) bool {
		if updates[i].Col == updates[j].Col {
			return updates[i].Ref < updates[j].Ref
		}
		return updates[i].Col < updates[j].Col
	})

	for _, update := range updates {
		_ = writeCell(encoder, cellName, update.Ref, nil, update.Value)
	}
}

func appendMissingRows(encoder *xml.Encoder, rowName, cellName xml.Name, pending map[string]cellUpdate, seenRows map[int]bool) {
	if rowName.Local == "" {
		rowName = xml.Name{Local: "row"}
	}
	if cellName.Local == "" {
		cellName = xml.Name{Local: "c"}
	}

	rows := make(map[int][]cellUpdate)
	for _, update := range pending {
		if seenRows[update.Row] {
			continue
		}
		rows[update.Row] = append(rows[update.Row], update)
	}

	rowNumbers := make([]int, 0, len(rows))
	for row := range rows {
		rowNumbers = append(rowNumbers, row)
	}
	sort.Ints(rowNumbers)

	for _, row := range rowNumbers {
		start := xml.StartElement{
			Name: rowName,
			Attr: []xml.Attr{{Name: xml.Name{Local: "r"}, Value: strconv.Itoa(row)}},
		}
		_ = encoder.EncodeToken(start)

		cells := rows[row]
		sort.Slice(cells, func(i, j int) bool {
			if cells[i].Col == cells[j].Col {
				return cells[i].Ref < cells[j].Ref
			}
			return cells[i].Col < cells[j].Col
		})
		for _, cell := range cells {
			_ = writeCell(encoder, cellName, cell.Ref, nil, cell.Value)
			delete(pending, cell.Ref)
		}

		_ = encoder.EncodeToken(xml.EndElement{Name: rowName})
	}
}

func writeCell(encoder *xml.Encoder, name xml.Name, cellRef string, attrs []xml.Attr, value CellValue) error {
	start := xml.StartElement{Name: name, Attr: buildCellAttrs(cellRef, attrs, value)}
	if err := encoder.EncodeToken(start); err != nil {
		return err
	}

	if value.String != nil {
		if err := writeInlineStr(encoder, *value.String); err != nil {
			return err
		}
	} else if value.Number != nil {
		if err := writeNumberValue(encoder, *value.Number); err != nil {
			return err
		}
	}

	if err := encoder.EncodeToken(xml.EndElement{Name: name}); err != nil {
		return err
	}
	return nil
}

func writeInlineStr(encoder *xml.Encoder, value string) error {
	if err := encoder.EncodeToken(xml.StartElement{Name: xml.Name{Local: "is"}}); err != nil {
		return err
	}
	if err := encoder.EncodeToken(xml.StartElement{Name: xml.Name{Local: "t"}}); err != nil {
		return err
	}
	if err := encoder.EncodeToken(xml.CharData([]byte(value))); err != nil {
		return err
	}
	if err := encoder.EncodeToken(xml.EndElement{Name: xml.Name{Local: "t"}}); err != nil {
		return err
	}
	if err := encoder.EncodeToken(xml.EndElement{Name: xml.Name{Local: "is"}}); err != nil {
		return err
	}
	return nil
}

func writeNumberValue(encoder *xml.Encoder, value float64) error {
	if err := encoder.EncodeToken(xml.StartElement{Name: xml.Name{Local: "v"}}); err != nil {
		return err
	}
	if err := encoder.EncodeToken(xml.CharData([]byte(formatNumber(value)))); err != nil {
		return err
	}
	if err := encoder.EncodeToken(xml.EndElement{Name: xml.Name{Local: "v"}}); err != nil {
		return err
	}
	return nil
}

func buildCellAttrs(cellRef string, attrs []xml.Attr, value CellValue) []xml.Attr {
	out := make([]xml.Attr, 0, len(attrs)+2)
	hasRef := false
	hasType := false

	for _, attr := range attrs {
		switch attr.Name.Local {
		case "r":
			attr.Value = cellRef
			hasRef = true
			out = append(out, attr)
		case "t":
			if value.String != nil {
				attr.Value = "inlineStr"
				out = append(out, attr)
				hasType = true
			}
		default:
			out = append(out, attr)
		}
	}

	if !hasRef {
		out = append(out, xml.Attr{Name: xml.Name{Local: "r"}, Value: cellRef})
	}
	if value.String != nil && !hasType {
		out = append(out, xml.Attr{Name: xml.Name{Local: "t"}, Value: "inlineStr"})
	}

	return out
}

func formatNumber(value float64) string {
	return strconv.FormatFloat(value, 'f', -1, 64)
}

func writeOverrideEntry(writer *zip.Writer, part *zip.File, data []byte) error {
	header := part.FileHeader
	if part.Flags&0x8 != 0 {
		return writeEntryWithDescriptor(writer, &header, data)
	}
	return writeRawEntry(writer, &header, data)
}

func writeEntryWithDescriptor(writer *zip.Writer, header *zip.FileHeader, data []byte) error {
	header.CRC32 = 0
	header.CompressedSize = 0
	header.UncompressedSize = 0
	header.CompressedSize64 = 0
	header.UncompressedSize64 = 0

	entry, err := writer.CreateHeader(header)
	if err != nil {
		return err
	}
	if len(data) == 0 {
		return nil
	}
	_, err = entry.Write(data)
	return err
}

func writeRawEntry(writer *zip.Writer, header *zip.FileHeader, data []byte) error {
	compressed, err := compressData(header.Method, data)
	if err != nil {
		return err
	}

	header.Flags &^= 0x8
	header.CRC32 = crc32.ChecksumIEEE(data)
	header.UncompressedSize64 = uint64(len(data))
	header.UncompressedSize = uint32(len(data))
	header.CompressedSize64 = uint64(len(compressed))
	header.CompressedSize = uint32(len(compressed))

	entry, err := writer.CreateRaw(header)
	if err != nil {
		return err
	}
	if len(compressed) == 0 {
		return nil
	}
	_, err = entry.Write(compressed)
	return err
}

func writeNewEntry(writer *zip.Writer, name string, data []byte) error {
	if strings.HasSuffix(name, "/") {
		_, err := writer.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Store})
		return err
	}

	header := zip.FileHeader{Name: name, Method: zip.Deflate}
	return writeRawEntry(writer, &header, data)
}

func compressData(method uint16, data []byte) ([]byte, error) {
	switch method {
	case zip.Store:
		return data, nil
	case zip.Deflate:
		var buf bytes.Buffer
		zw, err := flate.NewWriter(&buf, flate.DefaultCompression)
		if err != nil {
			return nil, err
		}
		if len(data) > 0 {
			if _, err := zw.Write(data); err != nil {
				_ = zw.Close()
				return nil, err
			}
		}
		if err := zw.Close(); err != nil {
			return nil, err
		}
		return buf.Bytes(), nil
	default:
		return nil, fmt.Errorf("unsupported compression method: %d", method)
	}
}

func colToIndex(col string) int {
	index := 0
	for i := 0; i < len(col); i++ {
		ch := col[i]
		if ch < 'A' || ch > 'Z' {
			return 0
		}
		index = index*26 + int(ch-'A'+1)
	}
	return index
}

func indexToCol(index int) string {
	if index <= 0 {
		return ""
	}
	var buf [8]byte
	pos := len(buf)
	for index > 0 {
		index--
		buf[pos-1] = byte('A' + (index % 26))
		pos--
		index /= 26
	}
	return string(buf[pos:])
}
