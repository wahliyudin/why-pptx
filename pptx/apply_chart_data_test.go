package pptx

import (
	"path/filepath"
	"testing"
)

func TestApplyChartDataUpdatesWorkbook(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	workbook := buildWorkbookWithValues(t, "Old1", "Old2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithCaches("Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3", []string{"Old1", "Old2"}, []string{"10", "20"}),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	opts := DefaultOptions()
	opts.Chart.CacheSync = false
	doc, err := OpenFile(inputPath, WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"NewA", "NewB"},
		"values:0":   {"100", "200"},
	}
	if err := doc.ApplyChartData(0, data); err != nil {
		t.Fatalf("ApplyChartData: %v", err)
	}
	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	updatedWorkbook := readEmbeddedWorkbook(t, outputPath, "ppt/embeddings/embeddedWorkbook1.xlsx")
	sheet := readSheetFromXLSX(t, updatedWorkbook, "xl/worksheets/sheet1.xml")

	typ, val, ok := readCellFromSheet(sheet, "A2")
	if !ok || typ != "inlineStr" || val != "NewA" {
		t.Fatalf("unexpected A2: type=%q val=%q ok=%v", typ, val, ok)
	}
	typ, val, ok = readCellFromSheet(sheet, "A3")
	if !ok || typ != "inlineStr" || val != "NewB" {
		t.Fatalf("unexpected A3: type=%q val=%q ok=%v", typ, val, ok)
	}
	typ, val, ok = readCellFromSheet(sheet, "B2")
	if !ok || typ != "" || val != "100" {
		t.Fatalf("unexpected B2: type=%q val=%q ok=%v", typ, val, ok)
	}
	typ, val, ok = readCellFromSheet(sheet, "B3")
	if !ok || typ != "" || val != "200" {
		t.Fatalf("unexpected B3: type=%q val=%q ok=%v", typ, val, ok)
	}
}

func TestApplyChartDataLengthMismatchStrict(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	chartXML := chartWithCaches("Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$4", []string{"Old1", "Old2"}, []string{"10", "20", "30"})
	workbook := buildWorkbookWithValues(t, "Old1", "Old2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartXML,
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"NewA", "NewB"},
		"values:0":   {"100", "200", "300"},
	}
	if err := doc.ApplyChartData(0, data); err == nil {
		t.Fatalf("expected length mismatch error")
	}

	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	outChart := readZipEntry(t, outputPath, "ppt/charts/chart1.xml")
	if string(outChart) != string(chartXML) {
		t.Fatalf("chart xml changed on mismatch")
	}

	updatedWorkbook := readEmbeddedWorkbook(t, outputPath, "ppt/embeddings/embeddedWorkbook1.xlsx")
	sheet := readSheetFromXLSX(t, updatedWorkbook, "xl/worksheets/sheet1.xml")
	typ, val, ok := readCellFromSheet(sheet, "A2")
	if !ok || typ != "inlineStr" || val != "Old1" {
		t.Fatalf("unexpected A2: type=%q val=%q ok=%v", typ, val, ok)
	}
	typ, val, ok = readCellFromSheet(sheet, "B2")
	if !ok || typ != "" || val != "10" {
		t.Fatalf("unexpected B2: type=%q val=%q ok=%v", typ, val, ok)
	}
}

func TestApplyChartDataLengthMismatchBestEffort(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	chartXML := chartWithCaches("Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$4", []string{"Old1", "Old2"}, []string{"10", "20", "30"})
	workbook := buildWorkbookWithValues(t, "Old1", "Old2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartXML,
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(inputPath, WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"NewA", "NewB"},
		"values:0":   {"100", "200", "300"},
	}
	if err := doc.ApplyChartData(0, data); err != nil {
		t.Fatalf("ApplyChartData: %v", err)
	}

	alerts := doc.AlertsByCode("CHART_DATA_LENGTH_MISMATCH")
	if len(alerts) != 1 {
		t.Fatalf("expected mismatch alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	outChart := readZipEntry(t, outputPath, "ppt/charts/chart1.xml")
	if string(outChart) != string(chartXML) {
		t.Fatalf("chart xml changed on mismatch")
	}

	updatedWorkbook := readEmbeddedWorkbook(t, outputPath, "ppt/embeddings/embeddedWorkbook1.xlsx")
	sheet := readSheetFromXLSX(t, updatedWorkbook, "xl/worksheets/sheet1.xml")
	typ, val, ok := readCellFromSheet(sheet, "A2")
	if !ok || typ != "inlineStr" || val != "Old1" {
		t.Fatalf("unexpected A2: type=%q val=%q ok=%v", typ, val, ok)
	}
	typ, val, ok = readCellFromSheet(sheet, "B2")
	if !ok || typ != "" || val != "10" {
		t.Fatalf("unexpected B2: type=%q val=%q ok=%v", typ, val, ok)
	}
}
