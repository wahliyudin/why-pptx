package postflight

import (
	"archive/zip"
	"bytes"
	"testing"

	"why-pptx/internal/overlaystage"
)

type alertRecord struct {
	code string
	ctx  map[string]string
}

func newValidator(parent overlaystage.Overlay, alerts *[]alertRecord) *PostflightValidator {
	doc := &Document{
		Overlay: parent,
		EmitAlert: func(code, message string, ctx map[string]string) {
			if alerts != nil {
				*alerts = append(*alerts, alertRecord{code: code, ctx: ctx})
			}
		},
	}
	return NewPostflightValidator(doc)
}

type memOverlay struct {
	baseline map[string][]byte
	overlay  map[string][]byte
}

func newMemOverlay(parts map[string][]byte) *memOverlay {
	baseline := make(map[string][]byte, len(parts))
	for name, data := range parts {
		copied := make([]byte, len(data))
		copy(copied, data)
		baseline[name] = copied
	}
	return &memOverlay{
		baseline: baseline,
		overlay:  make(map[string][]byte),
	}
}

func (m *memOverlay) Get(path string) ([]byte, error) {
	if data, ok := m.overlay[path]; ok {
		return append([]byte(nil), data...), nil
	}
	if data, ok := m.baseline[path]; ok {
		return append([]byte(nil), data...), nil
	}
	return nil, errNotFound(path)
}

func (m *memOverlay) Set(path string, content []byte) error {
	copied := make([]byte, len(content))
	copy(copied, content)
	m.overlay[path] = copied
	return nil
}

func (m *memOverlay) Has(path string) (bool, error) {
	if _, ok := m.overlay[path]; ok {
		return true, nil
	}
	_, ok := m.baseline[path]
	return ok, nil
}

func (m *memOverlay) ListEntries() ([]string, error) {
	names := make([]string, 0, len(m.baseline)+len(m.overlay))
	for name := range m.baseline {
		names = append(names, name)
	}
	for name := range m.overlay {
		names = append(names, name)
	}
	return names, nil
}

func (m *memOverlay) HasBaseline(path string) (bool, error) {
	_, ok := m.baseline[path]
	return ok, nil
}

type notFoundError struct {
	path string
}

func (e notFoundError) Error() string {
	return "not found: " + e.path
}

func errNotFound(path string) error {
	return notFoundError{path: path}
}

func TestPostflightMalformedChartXML(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("<c:chartSpace></c:chartSpace>"),
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)

	if err := stage.Set("ppt/charts/chart1.xml", []byte("<c:chartSpace><broken")); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{ChartPath: "ppt/charts/chart1.xml", Mode: ModeStrict}
	err := validator.ValidateChartStage(ctx, stage)
	if err == nil {
		t.Fatalf("expected malformed xml error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_XML_MALFORMED" {
		t.Fatalf("expected POSTFLIGHT_XML_MALFORMED alert, got %#v", alerts)
	}
}

func TestPostflightSharedStringsDetected(t *testing.T) {
	xlsx := buildXLSXWithSharedStrings(t)
	parent := newMemOverlay(map[string][]byte{
		"ppt/embeddings/embeddedWorkbook1.xlsx": xlsx,
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)

	if err := stage.Set("ppt/embeddings/embeddedWorkbook1.xlsx", xlsx); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx", Mode: ModeStrict}
	err := validator.ValidateChartStage(ctx, stage)
	if err == nil {
		t.Fatalf("expected sharedStrings error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED" {
		t.Fatalf("expected POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED alert, got %#v", alerts)
	}
}

func TestPostflightRelTargetMissing(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("<c:chartSpace></c:chartSpace>"),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="missing.png"/>
</Relationships>`),
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)

	if err := stage.Set("ppt/charts/chart1.xml", []byte("<c:chartSpace></c:chartSpace>")); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{ChartPath: "ppt/charts/chart1.xml", Mode: ModeStrict}
	err := validator.ValidateChartStage(ctx, stage)
	if err == nil {
		t.Fatalf("expected rel target missing error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_REL_TARGET_MISSING" {
		t.Fatalf("expected POSTFLIGHT_REL_TARGET_MISSING alert, got %#v", alerts)
	}
}

func TestPostflightRelTargetUsesStageView(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("<c:chartSpace></c:chartSpace>"),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`),
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)

	if err := stage.Set("ppt/media/image1.png", []byte("data")); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{ChartPath: "ppt/charts/chart1.xml", Mode: ModeStrict}
	if err := validator.checkRelationshipTargets(ctx, stage, "ppt/charts/chart1.xml"); err != nil {
		t.Fatalf("expected rel target check to pass, got %v", err)
	}
	if len(alerts) != 0 {
		t.Fatalf("unexpected alerts: %#v", alerts)
	}
}

func TestPostflightUnexpectedPartAdded(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("<c:chartSpace></c:chartSpace>"),
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)

	if err := stage.Set("ppt/charts/new.xml", []byte("<c:chartSpace></c:chartSpace>")); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{ChartPath: "ppt/charts/chart1.xml", Mode: ModeStrict}
	err := validator.ValidateChartStage(ctx, stage)
	if err == nil {
		t.Fatalf("expected unexpected part error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_UNEXPECTED_PART_ADDED" {
		t.Fatalf("expected POSTFLIGHT_UNEXPECTED_PART_ADDED alert, got %#v", alerts)
	}
}

func TestPostflightChartCachePtCountMismatch(t *testing.T) {
	chartXML := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`)

	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": chartXML,
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)
	if err := stage.Set("ppt/charts/chart1.xml", chartXML); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{ChartPath: "ppt/charts/chart1.xml", Mode: ModeStrict, CacheSyncEnabled: true}
	if err := validator.ValidateChartStage(ctx, stage); err == nil {
		t.Fatalf("expected chart cache error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_CHART_CACHE_INVALID" {
		t.Fatalf("expected POSTFLIGHT_CHART_CACHE_INVALID alert, got %#v", alerts)
	}
}

func TestPostflightChartCacheIdxGap(t *testing.T) {
	chartXML := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
                <c:pt idx="2"><c:v>2</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`)

	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": chartXML,
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)
	if err := stage.Set("ppt/charts/chart1.xml", chartXML); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{ChartPath: "ppt/charts/chart1.xml", Mode: ModeStrict, CacheSyncEnabled: true}
	if err := validator.ValidateChartStage(ctx, stage); err == nil {
		t.Fatalf("expected chart cache error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_CHART_CACHE_INVALID" {
		t.Fatalf("expected POSTFLIGHT_CHART_CACHE_INVALID alert, got %#v", alerts)
	}
}

func TestPostflightChartCacheInvalidNumeric(t *testing.T) {
	chartXML := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>abc</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`)

	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": chartXML,
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)
	if err := stage.Set("ppt/charts/chart1.xml", chartXML); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{
		ChartPath:            "ppt/charts/chart1.xml",
		Mode:                 ModeStrict,
		CacheSyncEnabled:     true,
		MissingNumericPolicy: 0,
	}
	if err := validator.ValidateChartStage(ctx, stage); err == nil {
		t.Fatalf("expected chart cache error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_CHART_CACHE_INVALID" {
		t.Fatalf("expected POSTFLIGHT_CHART_CACHE_INVALID alert, got %#v", alerts)
	}
}

func TestPostflightWorksheetSharedStringCellType(t *testing.T) {
	xlsx := buildXLSXWithSharedStringCell(t)
	parent := newMemOverlay(map[string][]byte{
		"ppt/embeddings/embeddedWorkbook1.xlsx": xlsx,
	})
	var alerts []alertRecord
	validator := newValidator(parent, &alerts)
	stage := overlaystage.NewStagingOverlay(parent)
	if err := stage.Set("ppt/embeddings/embeddedWorkbook1.xlsx", xlsx); err != nil {
		t.Fatalf("Set: %v", err)
	}

	ctx := ValidateContext{WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx", Mode: ModeStrict}
	if err := validator.ValidateChartStage(ctx, stage); err == nil {
		t.Fatalf("expected cell type mismatch error")
	}
	if len(alerts) != 1 || alerts[0].code != "POSTFLIGHT_XLSX_CELL_TYPE_MISMATCH" {
		t.Fatalf("expected POSTFLIGHT_XLSX_CELL_TYPE_MISMATCH alert, got %#v", alerts)
	}
}

func buildXLSXWithSharedStrings(t *testing.T) []byte {
	t.Helper()

	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)

	files := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
		"xl/sharedStrings.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>`),
	}

	for name, data := range files {
		entry, err := writer.Create(name)
		if err != nil {
			_ = writer.Close()
			t.Fatalf("Create: %v", err)
		}
		if _, err := entry.Write(data); err != nil {
			_ = writer.Close()
			t.Fatalf("Write: %v", err)
		}
	}

	if err := writer.Close(); err != nil {
		t.Fatalf("Close: %v", err)
	}
	return buf.Bytes()
}

func buildXLSXWithSharedStringCell(t *testing.T) []byte {
	t.Helper()

	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)

	files := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
		"xl/worksheets/sheet1.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>`),
	}

	for name, data := range files {
		entry, err := writer.Create(name)
		if err != nil {
			_ = writer.Close()
			t.Fatalf("Create: %v", err)
		}
		if _, err := entry.Write(data); err != nil {
			_ = writer.Close()
			t.Fatalf("Write: %v", err)
		}
	}

	if err := writer.Close(); err != nil {
		t.Fatalf("Close: %v", err)
	}
	return buf.Bytes()
}
