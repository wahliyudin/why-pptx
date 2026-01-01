package pptx

import (
	"reflect"
	"testing"
)

func TestExtractChartDataByPath_BarSimple(t *testing.T) {
	doc, err := OpenFile(fixturePath("bar_simple_embedded.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ExtractChartDataByPath: %v", err)
	}

	wantLabels := []string{"Old1", "Old2"}
	if data.Type != "bar" {
		t.Fatalf("expected bar chart, got %q", data.Type)
	}
	if !reflect.DeepEqual(data.Labels, wantLabels) {
		t.Fatalf("labels mismatch: got %v want %v", data.Labels, wantLabels)
	}
	if len(data.Series) != 1 {
		t.Fatalf("expected 1 series, got %d", len(data.Series))
	}
	if data.Series[0].Index != 0 {
		t.Fatalf("expected series index 0, got %d", data.Series[0].Index)
	}
	if data.Series[0].Name != "Series 1" {
		t.Fatalf("unexpected series name %q", data.Series[0].Name)
	}
	if !reflect.DeepEqual(data.Series[0].Data, []string{"10", "20"}) {
		t.Fatalf("series data mismatch: %v", data.Series[0].Data)
	}
	if data.Meta.ChartPath != "ppt/charts/chart1.xml" {
		t.Fatalf("unexpected chart path %q", data.Meta.ChartPath)
	}
	if data.Meta.WorkbookPath == "" || data.Meta.SlidePath == "" {
		t.Fatalf("missing metadata: %+v", data.Meta)
	}
	if data.Meta.Sheet != "Sheet1" {
		t.Fatalf("expected Sheet1, got %q", data.Meta.Sheet)
	}

	again, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ExtractChartDataByPath again: %v", err)
	}
	if !reflect.DeepEqual(data, again) {
		t.Fatalf("extract output not deterministic")
	}
}

func TestExtractChartDataByPath_LineMultiSeries(t *testing.T) {
	doc, err := OpenFile(fixturePath("line_multi_series_embedded.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ExtractChartDataByPath: %v", err)
	}

	if data.Type != "line" {
		t.Fatalf("expected line chart, got %q", data.Type)
	}
	wantLabels := []string{"Cat1", "Cat2", "Cat3"}
	if !reflect.DeepEqual(data.Labels, wantLabels) {
		t.Fatalf("labels mismatch: got %v want %v", data.Labels, wantLabels)
	}
	if len(data.Series) != 2 {
		t.Fatalf("expected 2 series, got %d", len(data.Series))
	}
	if data.Series[0].Index != 0 || data.Series[1].Index != 1 {
		t.Fatalf("unexpected series order: %+v", data.Series)
	}
	if data.Series[0].Name != "Series 1" || data.Series[1].Name != "Series 2" {
		t.Fatalf("unexpected series names: %+v", data.Series)
	}
	if !reflect.DeepEqual(data.Series[0].Data, []string{"1", "2", "3"}) {
		t.Fatalf("series 0 data mismatch: %v", data.Series[0].Data)
	}
	if !reflect.DeepEqual(data.Series[1].Data, []string{"4", "5", "6"}) {
		t.Fatalf("series 1 data mismatch: %v", data.Series[1].Data)
	}
}

func TestExtractChartDataByPath_PieSimple(t *testing.T) {
	doc, err := OpenFile(fixturePath("pie_simple_embedded.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ExtractChartDataByPath: %v", err)
	}

	if data.Type != "pie" {
		t.Fatalf("expected pie chart, got %q", data.Type)
	}
	if !reflect.DeepEqual(data.Labels, []string{"Slice1", "Slice2", "Slice3"}) {
		t.Fatalf("labels mismatch: %v", data.Labels)
	}
	if len(data.Series) != 1 {
		t.Fatalf("expected 1 series, got %d", len(data.Series))
	}
	if !reflect.DeepEqual(data.Series[0].Data, []string{"5", "15", "25"}) {
		t.Fatalf("series data mismatch: %v", data.Series[0].Data)
	}
}

func TestExtractChartDataByPath_AreaSimple(t *testing.T) {
	doc, err := OpenFile(fixturePath("area_simple_embedded.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ExtractChartDataByPath: %v", err)
	}

	if data.Type != "area" {
		t.Fatalf("expected area chart, got %q", data.Type)
	}
	if !reflect.DeepEqual(data.Labels, []string{"Cat1", "Cat2"}) {
		t.Fatalf("labels mismatch: %v", data.Labels)
	}
	if len(data.Series) != 1 {
		t.Fatalf("expected 1 series, got %d", len(data.Series))
	}
	if !reflect.DeepEqual(data.Series[0].Data, []string{"10", "20"}) {
		t.Fatalf("series data mismatch: %v", data.Series[0].Data)
	}
}

func TestExportChartByPathFormat_Pie(t *testing.T) {
	doc, err := OpenFile(fixturePath("pie_simple_embedded.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	payload, err := doc.ExportChartByPathFormat("ppt/charts/chart1.xml", ExportChartJS)
	if err != nil {
		t.Fatalf("ExportChartByPathFormat: %v", err)
	}
	if payload.Data["type"] != "pie" {
		t.Fatalf("unexpected chartjs type: %#v", payload.Data["type"])
	}
}

func TestExportChartByPathFormat_Area(t *testing.T) {
	doc, err := OpenFile(fixturePath("area_simple_embedded.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	payload, err := doc.ExportChartByPathFormat("ppt/charts/chart1.xml", ExportChartJS)
	if err != nil {
		t.Fatalf("ExportChartByPathFormat: %v", err)
	}
	if payload.Data["type"] != "line" {
		t.Fatalf("expected line type for area export, got %#v", payload.Data["type"])
	}
}

func TestExtractAllCharts_LinkedWorkbook_BestEffort(t *testing.T) {
	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(fixturePath("linked_workbook_chart.pptx"), WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.ExtractAllCharts()
	if err != nil {
		t.Fatalf("ExtractAllCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected no charts, got %d", len(charts))
	}

	alerts := doc.AlertsByCode("CHART_LINKED_WORKBOOK")
	if len(alerts) != 1 {
		t.Fatalf("expected CHART_LINKED_WORKBOOK alert, got %d", len(alerts))
	}
}

func TestExtractAllCharts_LinkedWorkbook_Strict(t *testing.T) {
	doc, err := OpenFile(fixturePath("linked_workbook_chart.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if _, err := doc.ExtractAllCharts(); err == nil {
		t.Fatalf("expected ExtractAllCharts error")
	}
}

func TestExtractAllCharts_SharedStrings_BestEffort(t *testing.T) {
	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(fixturePath("xlsx_sharedStrings_present.pptx"), WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.ExtractAllCharts()
	if err != nil {
		t.Fatalf("ExtractAllCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected no charts, got %d", len(charts))
	}

	alerts := doc.AlertsByCode("EXTRACT_SHAREDSTRINGS_UNSUPPORTED")
	if len(alerts) != 1 {
		t.Fatalf("expected EXTRACT_SHAREDSTRINGS_UNSUPPORTED alert, got %d", len(alerts))
	}
}

func TestExtractChartDataByPath_SharedStrings_Strict(t *testing.T) {
	doc, err := OpenFile(fixturePath("xlsx_sharedStrings_present.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if _, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml"); err == nil {
		t.Fatalf("expected ExtractChartDataByPath error")
	}
}

func TestExtractChartDataByPath_PieMultiSeries(t *testing.T) {
	dir := t.TempDir()
	path := dir + "/pie_multi_series.pptx"

	workbookBytes := buildTestXLSX(t)
	chartXML := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$4:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$4:$B$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`)

	parts := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>`),
		"ppt/slides/slide1.xml":                 []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels":      []byte(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`),
		"ppt/charts/chart1.xml":                 chartXML,
		"ppt/charts/_rels/chart1.xml.rels":      []byte(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/></Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbookBytes,
	}
	if err := writeZipFile(path, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	if _, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml"); err == nil {
		t.Fatalf("expected ExtractChartDataByPath error")
	}
}

func TestDetectSharedStringsPart(t *testing.T) {
	parts := baseXLSXParts(t)
	parts["xl/sharedStrings.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>`)
	data := writeZipBytes(t, parts)

	found, partPath, cellRef, err := detectSharedStrings(data)
	if err != nil {
		t.Fatalf("detectSharedStrings: %v", err)
	}
	if !found {
		t.Fatalf("expected sharedStrings detection")
	}
	if partPath != "xl/sharedStrings.xml" {
		t.Fatalf("unexpected part path %q", partPath)
	}
	if cellRef != "" {
		t.Fatalf("unexpected cell ref %q", cellRef)
	}
}

func TestDetectSharedStringsCell(t *testing.T) {
	parts := baseXLSXParts(t)
	parts["xl/worksheets/sheet1.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`)
	data := writeZipBytes(t, parts)

	found, partPath, cellRef, err := detectSharedStrings(data)
	if err != nil {
		t.Fatalf("detectSharedStrings: %v", err)
	}
	if !found {
		t.Fatalf("expected shared string cell detection")
	}
	if partPath != "xl/worksheets/sheet1.xml" {
		t.Fatalf("unexpected sheet path %q", partPath)
	}
	if cellRef != "A1" {
		t.Fatalf("unexpected cell ref %q", cellRef)
	}
}

func TestExpandRangeCells(t *testing.T) {
	cells, err := expandRangeCells("A2", "A4")
	if err != nil {
		t.Fatalf("expandRangeCells column: %v", err)
	}
	if !reflect.DeepEqual(cells, []string{"A2", "A3", "A4"}) {
		t.Fatalf("column range mismatch: %v", cells)
	}

	cells, err = expandRangeCells("A2", "C2")
	if err != nil {
		t.Fatalf("expandRangeCells row: %v", err)
	}
	if !reflect.DeepEqual(cells, []string{"A2", "B2", "C2"}) {
		t.Fatalf("row range mismatch: %v", cells)
	}

	if _, err := expandRangeCells("A1", "B2"); err == nil {
		t.Fatalf("expected 2D range error")
	}
}

func baseXLSXParts(t *testing.T) map[string][]byte {
	t.Helper()

	return map[string][]byte{
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
}
