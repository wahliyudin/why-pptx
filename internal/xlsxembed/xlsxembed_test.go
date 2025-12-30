package xlsxembed

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"sort"
	"testing"
)

func TestSetCellNumericExisting(t *testing.T) {
	data := buildTestXLSX(t)
	wb, err := Open(data)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	value := 42.0
	if err := wb.SetCell("Sheet1", "A1", CellValue{Number: &value}); err != nil {
		t.Fatalf("SetCell: %v", err)
	}

	out, err := wb.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}

	sheetData := readSheet(t, out, "xl/worksheets/sheet1.xml")
	typ, val, ok := readCell(sheetData, "A1")
	if !ok {
		t.Fatalf("expected cell A1")
	}
	if typ != "" || val != "42" {
		t.Fatalf("unexpected cell value: type=%q val=%q", typ, val)
	}
}

func TestSetCellCreatesMissing(t *testing.T) {
	data := buildTestXLSX(t)
	wb, err := Open(data)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	value := 3.5
	if err := wb.SetCell("Sheet1", "B3", CellValue{Number: &value}); err != nil {
		t.Fatalf("SetCell: %v", err)
	}

	out, err := wb.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}

	sheetData := readSheet(t, out, "xl/worksheets/sheet1.xml")
	typ, val, ok := readCell(sheetData, "B3")
	if !ok {
		t.Fatalf("expected cell B3")
	}
	if typ != "" || val != "3.5" {
		t.Fatalf("unexpected cell value: type=%q val=%q", typ, val)
	}
}

func TestSetCellInlineStr(t *testing.T) {
	data := buildTestXLSX(t)
	wb, err := Open(data)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	text := "hello"
	if err := wb.SetCell("Sheet1", "C2", CellValue{String: &text}); err != nil {
		t.Fatalf("SetCell: %v", err)
	}

	out, err := wb.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}

	sheetData := readSheet(t, out, "xl/worksheets/sheet1.xml")
	typ, val, ok := readCell(sheetData, "C2")
	if !ok {
		t.Fatalf("expected cell C2")
	}
	if typ != "inlineStr" || val != "hello" {
		t.Fatalf("unexpected cell value: type=%q val=%q", typ, val)
	}
}

func TestSetCellUnicodeSheetName(t *testing.T) {
	data := buildTestXLSX(t)
	wb, err := Open(data)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	value := 7.0
	if err := wb.SetCell("Data 📈", "A1", CellValue{Number: &value}); err != nil {
		t.Fatalf("SetCell: %v", err)
	}

	out, err := wb.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}

	sheetData := readSheet(t, out, "xl/worksheets/sheet2.xml")
	typ, val, ok := readCell(sheetData, "A1")
	if !ok {
		t.Fatalf("expected cell A1")
	}
	if typ != "" || val != "7" {
		t.Fatalf("unexpected cell value: type=%q val=%q", typ, val)
	}
}

func TestPreservesSheetExtras(t *testing.T) {
	data := buildTestXLSXWithExtras(t)
	wb, err := Open(data)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	value := 5.0
	if err := wb.SetCell("Sheet1", "A1", CellValue{Number: &value}); err != nil {
		t.Fatalf("SetCell: %v", err)
	}

	out, err := wb.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}

	sheetData := readSheet(t, out, "xl/worksheets/sheet1.xml")
	if !bytes.Contains(sheetData, []byte("<cols")) {
		t.Fatalf("expected cols element preserved")
	}
	if !bytes.Contains(sheetData, []byte("<mergeCells")) {
		t.Fatalf("expected mergeCells element preserved")
	}

	typ, val, ok := readCell(sheetData, "A1")
	if !ok {
		t.Fatalf("expected cell A1")
	}
	if typ != "" || val != "5" {
		t.Fatalf("unexpected cell value: type=%q val=%q", typ, val)
	}
	typ, val, ok = readCell(sheetData, "B2")
	if !ok {
		t.Fatalf("expected cell B2")
	}
	if typ != "" || val != "2" {
		t.Fatalf("unexpected cell value for B2: type=%q val=%q", typ, val)
	}
}

func TestGetRangeValuesColumn(t *testing.T) {
	data := buildTestXLSX(t)
	wb, err := Open(data)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	values, err := wb.GetRangeValues("Sheet1", "A1", "A2")
	if err != nil {
		t.Fatalf("GetRangeValues: %v", err)
	}
	if len(values) != 2 {
		t.Fatalf("expected 2 values, got %d", len(values))
	}
	if values[0] != "1" || values[1] != "" {
		t.Fatalf("unexpected values: %#v", values)
	}
}

func TestGetRangeValuesRow(t *testing.T) {
	data := buildTestXLSX(t)
	wb, err := Open(data)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}

	text := "hello"
	if err := wb.SetCell("Sheet1", "B1", CellValue{String: &text}); err != nil {
		t.Fatalf("SetCell: %v", err)
	}
	updated, err := wb.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}
	wb, err = Open(updated)
	if err != nil {
		t.Fatalf("Open updated: %v", err)
	}

	values, err := wb.GetRangeValues("Sheet1", "A1", "B1")
	if err != nil {
		t.Fatalf("GetRangeValues: %v", err)
	}
	if len(values) != 2 {
		t.Fatalf("expected 2 values, got %d", len(values))
	}
	if values[0] != "1" || values[1] != "hello" {
		t.Fatalf("unexpected values: %#v", values)
	}
}

func buildTestXLSX(t *testing.T) []byte {
	t.Helper()

	parts := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
		"xl/workbook.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Data 📈" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
		"xl/worksheets/sheet1.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`),
		"xl/worksheets/sheet2.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
  </sheetData>
</worksheet>`),
	}

	return writeZip(t, parts)
}

func buildTestXLSXWithExtras(t *testing.T) []byte {
	t.Helper()

	parts := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
		"xl/workbook.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
		"xl/worksheets/sheet1.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>
    <col min="1" max="1" width="10" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="B2"><v>2</v></c>
    </row>
  </sheetData>
  <mergeCells count="1">
    <mergeCell ref="A1:B1"/>
  </mergeCells>
</worksheet>`),
	}

	return writeZip(t, parts)
}

func writeZip(t *testing.T, parts map[string][]byte) []byte {
	t.Helper()

	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)

	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)

	for _, name := range names {
		entry, err := writer.Create(name)
		if err != nil {
			_ = writer.Close()
			t.Fatalf("Create: %v", err)
		}
		if _, err := entry.Write(parts[name]); err != nil {
			_ = writer.Close()
			t.Fatalf("Write: %v", err)
		}
	}

	if err := writer.Close(); err != nil {
		t.Fatalf("Close: %v", err)
	}

	return buf.Bytes()
}

func readSheet(t *testing.T, data []byte, path string) []byte {
	t.Helper()

	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("NewReader: %v", err)
	}

	for _, part := range reader.File {
		if part.Name == path {
			rc, err := part.Open()
			if err != nil {
				t.Fatalf("Open: %v", err)
			}
			defer rc.Close()
			out, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("Read: %v", err)
			}
			return out
		}
	}

	t.Fatalf("sheet %q not found", path)
	return nil
}

func readCell(data []byte, ref string) (string, string, bool) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	inCell := false
	cellType := ""
	target := ""
	inValue := false
	var value bytes.Buffer

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", "", false
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "c":
				cellType = ""
				target = ""
				for _, attr := range tok.Attr {
					if attr.Name.Local == "r" {
						target = attr.Value
					}
					if attr.Name.Local == "t" {
						cellType = attr.Value
					}
				}
				if target == ref {
					inCell = true
				}
			case "v", "t":
				if inCell {
					inValue = true
					value.Reset()
				}
			}
		case xml.EndElement:
			if tok.Name.Local == "c" && inCell {
				return cellType, value.String(), true
			}
			if tok.Name.Local == "v" || tok.Name.Local == "t" {
				inValue = false
			}
		case xml.CharData:
			if inCell && inValue {
				value.Write([]byte(tok))
			}
		}
	}

	return "", "", false
}
