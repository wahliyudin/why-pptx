package pptx

import (
	"archive/zip"
	"os"
	"path/filepath"
	"testing"
)

func TestDiscoverEmbeddedCharts(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte("<chart/>"),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/book1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/book1.xlsx": []byte("workbook"),
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.DiscoverEmbeddedCharts()
	if err != nil {
		t.Fatalf("DiscoverEmbeddedCharts: %v", err)
	}
	if len(charts) != 1 {
		t.Fatalf("expected 1 chart, got %d", len(charts))
	}
	if charts[0].WorkbookPath != "ppt/embeddings/book1.xlsx" {
		t.Fatalf("unexpected workbook path: %q", charts[0].WorkbookPath)
	}
	if len(doc.Alerts()) != 0 {
		t.Fatalf("expected no alerts, got %d", len(doc.Alerts()))
	}
}

func TestDiscoverLinkedWorkbookAlert(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte("<chart/>"),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="https://example.com/book.xlsx" TargetMode="External"/>
</Relationships>`),
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.DiscoverEmbeddedCharts()
	if err != nil {
		t.Fatalf("DiscoverEmbeddedCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected 0 charts, got %d", len(charts))
	}

	alerts := doc.Alerts()
	if len(alerts) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(alerts))
	}
	if alerts[0].Code != "CHART_LINKED_WORKBOOK" {
		t.Fatalf("unexpected alert code: %q", alerts[0].Code)
	}
	if alerts[0].Context["target"] != "https://example.com/book.xlsx" {
		t.Fatalf("unexpected target: %q", alerts[0].Context["target"])
	}
}

func TestDiscoverChartTypeSuffixMatch(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/not-a-chart-related" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte("<chart/>"),
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.DiscoverEmbeddedCharts()
	if err != nil {
		t.Fatalf("DiscoverEmbeddedCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected 0 charts, got %d", len(charts))
	}
	if len(doc.Alerts()) != 0 {
		t.Fatalf("expected no alerts, got %d", len(doc.Alerts()))
	}
}

func TestDiscoverMissingChartRelsAlert(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte("<chart/>"),
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.DiscoverEmbeddedCharts()
	if err != nil {
		t.Fatalf("DiscoverEmbeddedCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected 0 charts, got %d", len(charts))
	}

	alerts := doc.Alerts()
	if len(alerts) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(alerts))
	}
	if alerts[0].Code != "CHART_RELS_MISSING" {
		t.Fatalf("unexpected alert code: %q", alerts[0].Code)
	}
	if alerts[0].Context["rels_path"] != "ppt/charts/_rels/chart1.xml.rels" {
		t.Fatalf("unexpected rels_path: %q", alerts[0].Context["rels_path"])
	}
}

func TestDiscoverWorkbookNotFoundAlert(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte("<chart/>"),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/style" Target="../style1.xml"/>
</Relationships>`),
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.DiscoverEmbeddedCharts()
	if err != nil {
		t.Fatalf("DiscoverEmbeddedCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected 0 charts, got %d", len(charts))
	}

	alerts := doc.Alerts()
	if len(alerts) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(alerts))
	}
	if alerts[0].Code != "CHART_WORKBOOK_NOT_FOUND" {
		t.Fatalf("unexpected alert code: %q", alerts[0].Code)
	}
}

func TestDiscoverUnsupportedWorkbookTargetAlert(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte("<chart/>"),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../media/book1.xlsx"/>
</Relationships>`),
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.DiscoverEmbeddedCharts()
	if err != nil {
		t.Fatalf("DiscoverEmbeddedCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected 0 charts, got %d", len(charts))
	}

	alerts := doc.Alerts()
	if len(alerts) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(alerts))
	}
	if alerts[0].Code != "CHART_WORKBOOK_UNSUPPORTED_TARGET" {
		t.Fatalf("unexpected alert code: %q", alerts[0].Code)
	}
	if alerts[0].Context["target"] != "ppt/media/book1.xlsx" {
		t.Fatalf("unexpected target: %q", alerts[0].Context["target"])
	}
}

func TestGetChartDependencies(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/book1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/book1.xlsx": []byte{},
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	deps, err := doc.GetChartDependencies()
	if err != nil {
		t.Fatalf("GetChartDependencies: %v", err)
	}
	if len(deps) != 1 {
		t.Fatalf("expected 1 dependency, got %d", len(deps))
	}
	if deps[0].WorkbookPath != "ppt/embeddings/book1.xlsx" {
		t.Fatalf("unexpected workbook path: %q", deps[0].WorkbookPath)
	}
	if deps[0].ChartType != "bar" {
		t.Fatalf("unexpected chart type: %q", deps[0].ChartType)
	}
	if len(deps[0].Ranges) != 2 {
		t.Fatalf("expected 2 ranges, got %d", len(deps[0].Ranges))
	}
	if deps[0].Ranges[0].Sheet != "Sheet1" {
		t.Fatalf("unexpected sheet: %q", deps[0].Ranges[0].Sheet)
	}
}

func TestGetChartDependenciesStrictInvalidFormula(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val><c:numRef><c:f>Sheet1!A1:A2</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`),
		"ppt/charts/chart2.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val><c:numRef><c:f>Sheet1A2</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/book1.xlsx"/>
</Relationships>`),
		"ppt/charts/_rels/chart2.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/book2.xlsx"/>
</Relationships>`),
		"ppt/embeddings/book1.xlsx": []byte{},
		"ppt/embeddings/book2.xlsx": []byte{},
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if _, err := doc.GetChartDependencies(); err == nil {
		t.Fatalf("expected error for invalid formula in strict mode")
	}
}

func TestGetChartDependenciesBestEffortInvalidFormula(t *testing.T) {
	dir := t.TempDir()
	path := filepath.Join(dir, "input.pptx")

	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val><c:numRef><c:f>Sheet1!A1:A2</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`),
		"ppt/charts/chart2.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val><c:numRef><c:f>Sheet1A2</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/book1.xlsx"/>
</Relationships>`),
		"ppt/charts/_rels/chart2.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/book2.xlsx"/>
</Relationships>`),
		"ppt/embeddings/book1.xlsx": []byte{},
		"ppt/embeddings/book2.xlsx": []byte{},
	}

	if err := writeZip(path, parts); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	doc, err := OpenFile(path, WithErrorMode(BestEffort))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	deps, err := doc.GetChartDependencies()
	if err != nil {
		t.Fatalf("GetChartDependencies: %v", err)
	}
	if len(deps) != 1 {
		t.Fatalf("expected 1 dependency, got %d", len(deps))
	}
	if deps[0].ChartPath != "ppt/charts/chart1.xml" {
		t.Fatalf("unexpected chart path: %q", deps[0].ChartPath)
	}

	alerts := doc.Alerts()
	if len(alerts) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(alerts))
	}
	if alerts[0].Code != "CHART_DEPENDENCIES_PARSE_FAILED" {
		t.Fatalf("unexpected alert code: %q", alerts[0].Code)
	}
	if alerts[0].Context["chart"] != "ppt/charts/chart2.xml" {
		t.Fatalf("unexpected chart context: %q", alerts[0].Context["chart"])
	}
	if alerts[0].Context["error"] == "" {
		t.Fatalf("expected error context")
	}
}

func writeZip(path string, parts map[string][]byte) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := zip.NewWriter(file)
	for name, data := range parts {
		entry, err := writer.Create(name)
		if err != nil {
			_ = writer.Close()
			return err
		}
		if _, err := entry.Write(data); err != nil {
			_ = writer.Close()
			return err
		}
	}

	if err := writer.Close(); err != nil {
		return err
	}

	return nil
}
