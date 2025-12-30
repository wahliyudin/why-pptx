package pptx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"sort"
	"testing"

	"why-pptx/internal/xlsxembed"
)

func TestSetWorkbookCellsIntegration(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	workbookBytes := buildTestXLSX(t)
	pptxParts := map[string][]byte{
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbookBytes,
	}
	if err := writeZipFile(inputPath, pptxParts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	value := 9.0
	if err := doc.SetWorkbookCells([]CellUpdate{{
		WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx",
		Sheet:        "Sheet1",
		Cell:         "A1",
		Value:        Num(value),
	}}); err != nil {
		t.Fatalf("SetWorkbookCells: %v", err)
	}

	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	updatedWorkbook := readEmbeddedWorkbook(t, outputPath, "ppt/embeddings/embeddedWorkbook1.xlsx")
	wb, err := xlsxembed.Open(updatedWorkbook)
	if err != nil {
		t.Fatalf("xlsxembed.Open: %v", err)
	}

	out, err := wb.Save()
	if err != nil {
		t.Fatalf("Save: %v", err)
	}
	sheetData := readSheetFromXLSX(t, out, "xl/worksheets/sheet1.xml")
	typ, val, ok := readCellFromSheet(sheetData, "A1")
	if !ok {
		t.Fatalf("expected cell A1")
	}
	if typ != "" || val != "9" {
		t.Fatalf("unexpected cell value: type=%q val=%q", typ, val)
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
  </sheets>
</workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
		"xl/worksheets/sheet1.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`),
	}

	return writeZipBytes(t, parts)
}

func writeZipFile(path string, parts map[string][]byte) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := zip.NewWriter(file)
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)

	for _, name := range names {
		entry, err := writer.Create(name)
		if err != nil {
			_ = writer.Close()
			return err
		}
		if _, err := entry.Write(parts[name]); err != nil {
			_ = writer.Close()
			return err
		}
	}

	if err := writer.Close(); err != nil {
		return err
	}
	return nil
}

func writeZipBytes(t *testing.T, parts map[string][]byte) []byte {
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

func readEmbeddedWorkbook(t *testing.T, path, entryName string) []byte {
	t.Helper()

	file, err := os.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer file.Close()

	info, err := file.Stat()
	if err != nil {
		t.Fatalf("Stat: %v", err)
	}

	reader, err := zip.NewReader(file, info.Size())
	if err != nil {
		t.Fatalf("NewReader: %v", err)
	}

	for _, part := range reader.File {
		if part.Name == entryName {
			rc, err := part.Open()
			if err != nil {
				t.Fatalf("Open entry: %v", err)
			}
			defer rc.Close()
			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("Read entry: %v", err)
			}
			return data
		}
	}

	t.Fatalf("entry %q not found", entryName)
	return nil
}

func readSheetFromXLSX(t *testing.T, data []byte, path string) []byte {
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

func readCellFromSheet(data []byte, ref string) (string, string, bool) {
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
