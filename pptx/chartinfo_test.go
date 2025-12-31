package pptx

import (
	"fmt"
	"path/filepath"
	"strings"
	"testing"
)

func TestListChartsBasic(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")

	workbook := buildWorkbookWithValues(t, "Old1", "Old2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>`),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithTitleAndSeries("Revenue", 2),
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

	charts, err := doc.ListCharts()
	if err != nil {
		t.Fatalf("ListCharts: %v", err)
	}
	if len(charts) != 1 {
		t.Fatalf("expected 1 chart, got %d", len(charts))
	}

	info := charts[0]
	if info.Index != 0 {
		t.Fatalf("unexpected index: %d", info.Index)
	}
	if info.ChartPath != "ppt/charts/chart1.xml" {
		t.Fatalf("unexpected chart path: %q", info.ChartPath)
	}
	if info.WorkbookPath != "ppt/embeddings/embeddedWorkbook1.xlsx" {
		t.Fatalf("unexpected workbook path: %q", info.WorkbookPath)
	}
	if info.ChartType != "bar" {
		t.Fatalf("unexpected chart type: %q", info.ChartType)
	}
	if info.SeriesCount != 2 {
		t.Fatalf("unexpected series count: %d", info.SeriesCount)
	}
	if info.Title != "Revenue" {
		t.Fatalf("unexpected title: %q", info.Title)
	}
}

func TestApplyChartDataByNameTitleMatch(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	workbook := buildWorkbookWithValues(t, "Old1", "Old2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>`),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithTitleAndRanges("Revenue", "Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3"),
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
	if err := doc.ApplyChartDataByName("Revenue", data); err != nil {
		t.Fatalf("ApplyChartDataByName: %v", err)
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
	typ, val, ok = readCellFromSheet(sheet, "B2")
	if !ok || typ != "" || val != "100" {
		t.Fatalf("unexpected B2: type=%q val=%q ok=%v", typ, val, ok)
	}
}

func TestApplyChartDataByNameAmbiguous(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	workbook := buildWorkbookWithValues(t, "Old1", "Old2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>`),
		"ppt/slides/slide2.xml": []byte(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>`),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/slides/_rels/slide2.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithTitleAndRanges("Revenue", "Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3"),
		"ppt/charts/chart2.xml": chartWithTitleAndRanges("Revenue", "Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3"),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/charts/_rels/chart2.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"NewA", "NewB"},
		"values:0":   {"100", "200"},
	}

	doc, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	if err := doc.ApplyChartDataByName("Revenue", data); err == nil {
		t.Fatalf("expected ambiguous name error")
	}

	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err = OpenFile(inputPath, WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	if err := doc.ApplyChartDataByName("Revenue", data); err == nil {
		t.Fatalf("expected ambiguous name error")
	}
	alerts := doc.AlertsByCode("CHART_NAME_AMBIGUOUS")
	if len(alerts) != 1 {
		t.Fatalf("expected ambiguous alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}
	updatedWorkbook := readEmbeddedWorkbook(t, outputPath, "ppt/embeddings/embeddedWorkbook1.xlsx")
	sheet := readSheetFromXLSX(t, updatedWorkbook, "xl/worksheets/sheet1.xml")
	typ, val, ok := readCellFromSheet(sheet, "A2")
	if !ok || typ != "inlineStr" || val != "Old1" {
		t.Fatalf("unexpected A2: type=%q val=%q ok=%v", typ, val, ok)
	}
}

func TestListChartsBestEffortMalformedChart(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")

	workbook := buildWorkbookWithValues(t, "Old1", "Old2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>`),
		"ppt/slides/slide2.xml": []byte(`<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"></p:sld>`),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/slides/_rels/slide2.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithTitleAndSeries("Revenue", 1),
		"ppt/charts/chart2.xml": []byte(`<c:chartSpace><broken`),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/charts/_rels/chart2.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
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

	charts, err := doc.ListCharts()
	if err != nil {
		t.Fatalf("ListCharts: %v", err)
	}
	if len(charts) != 2 {
		t.Fatalf("expected 2 charts, got %d", len(charts))
	}
	if charts[0].ChartType != "bar" {
		t.Fatalf("unexpected chart type: %q", charts[0].ChartType)
	}
	if charts[1].ChartType != "unknown" {
		t.Fatalf("expected unknown chart type, got %q", charts[1].ChartType)
	}
	alerts := doc.AlertsByCode("CHART_INFO_PARSE_FAILED")
	if len(alerts) != 1 {
		t.Fatalf("expected parse failed alert, got %d", len(alerts))
	}
}

func chartWithTitleAndSeries(title string, seriesCount int) []byte {
	var titleXML string
	if title != "" {
		titleXML = `<c:title><c:tx><c:rich><a:p><a:r><a:t>` + title + `</a:t></a:r></a:p></c:rich></c:tx></c:title>`
	}

	var series strings.Builder
	for i := 0; i < seriesCount; i++ {
		series.WriteString("<c:ser></c:ser>")
	}

	return []byte(fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    %s
    <c:plotArea>
      <c:barChart>%s</c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`, titleXML, series.String()))
}

func chartWithTitleAndRanges(title, catFormula, valFormula string) []byte {
	titleXML := ""
	if title != "" {
		titleXML = `<c:title><c:tx><c:rich><a:p><a:r><a:t>` + title + `</a:t></a:r></a:p></c:rich></c:tx></c:title>`
	}

	return []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    ` + titleXML + `
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>` + catFormula + `</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>` + valFormula + `</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`)
}
