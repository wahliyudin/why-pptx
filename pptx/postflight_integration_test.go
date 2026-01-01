package pptx

import (
	"bytes"
	"path/filepath"
	"testing"

	"why-pptx/internal/overlaystage"
	"why-pptx/internal/postflight"
)

func TestPostflightBestEffortDiscard(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	chartPath := "ppt/charts/chart1.xml"
	original := []byte("<c:chartSpace></c:chartSpace>")
	parts := map[string][]byte{
		chartPath: original,
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

	ctx := postflight.ValidateContext{
		ChartPath: chartPath,
		Mode:      postflight.ModeBestEffort,
	}
	err = doc.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		return stage.Set(chartPath, []byte("<c:chartSpace><broken"))
	})
	if err == nil {
		t.Fatalf("expected postflight error")
	}

	alerts := doc.AlertsByCode("POSTFLIGHT_XML_MALFORMED")
	if len(alerts) != 1 {
		t.Fatalf("expected POSTFLIGHT_XML_MALFORMED alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	outChart := readZipEntry(t, outputPath, chartPath)
	if !bytes.Equal(outChart, original) {
		t.Fatalf("chart xml changed after discard")
	}
}

func TestPostflightStrictAbort(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	chartPath := "ppt/charts/chart1.xml"
	original := []byte("<c:chartSpace></c:chartSpace>")
	parts := map[string][]byte{
		chartPath: original,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	ctx := postflight.ValidateContext{
		ChartPath: chartPath,
		Mode:      postflight.ModeStrict,
	}
	err = doc.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		return stage.Set(chartPath, []byte("<c:chartSpace><broken"))
	})
	if err == nil {
		t.Fatalf("expected postflight error")
	}

	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	outChart := readZipEntry(t, outputPath, chartPath)
	if !bytes.Equal(outChart, original) {
		t.Fatalf("chart xml changed after discard")
	}
}

func TestPostflightBestEffortDiscardChartCache(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	chartPath := "ppt/charts/chart1.xml"
	original := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`)
	invalid := []byte(`<?xml version="1.0" encoding="UTF-8"?>
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
	parts := map[string][]byte{
		chartPath: original,
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

	ctx := postflight.ValidateContext{
		ChartPath:        chartPath,
		Mode:             postflight.ModeBestEffort,
		CacheSyncEnabled: true,
	}
	err = doc.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		return stage.Set(chartPath, invalid)
	})
	if err == nil {
		t.Fatalf("expected postflight error")
	}

	alerts := doc.AlertsByCode("POSTFLIGHT_CHART_CACHE_INVALID")
	if len(alerts) != 1 {
		t.Fatalf("expected POSTFLIGHT_CHART_CACHE_INVALID alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}
	outChart := readZipEntry(t, outputPath, chartPath)
	if !bytes.Equal(outChart, original) {
		t.Fatalf("chart xml changed after discard")
	}
}

func TestPostflightStrictChartCacheAbort(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")

	chartPath := "ppt/charts/chart1.xml"
	original := []byte(`<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`)
	invalid := []byte(`<?xml version="1.0" encoding="UTF-8"?>
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
	parts := map[string][]byte{
		chartPath: original,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	ctx := postflight.ValidateContext{
		ChartPath:        chartPath,
		Mode:             postflight.ModeStrict,
		CacheSyncEnabled: true,
	}
	err = doc.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		return stage.Set(chartPath, invalid)
	})
	if err == nil {
		t.Fatalf("expected postflight error")
	}
}

func TestPostflightInlineStrViolation(t *testing.T) {
	cases := []struct {
		name string
		mode ErrorMode
	}{
		{name: "best-effort", mode: BestEffort},
		{name: "strict", mode: Strict},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			inputPath := filepath.Join(dir, "input.pptx")
			outputPath := filepath.Join(dir, "output.pptx")

			workbookPath := "ppt/embeddings/embeddedWorkbook1.xlsx"
			workbook := buildXLSXWithSharedStringCell(t)
			parts := map[string][]byte{
				workbookPath: workbook,
			}

			if err := writeZipFile(inputPath, parts); err != nil {
				t.Fatalf("writeZipFile: %v", err)
			}

			opts := DefaultOptions()
			opts.Mode = tc.mode
			doc, err := OpenFile(inputPath, WithOptions(opts))
			if err != nil {
				t.Fatalf("OpenFile: %v", err)
			}

			ctx := postflight.ValidateContext{
				WorkbookPath: workbookPath,
			}
			if tc.mode == BestEffort {
				ctx.Mode = postflight.ModeBestEffort
			} else {
				ctx.Mode = postflight.ModeStrict
			}
			err = doc.withChartStage(ctx, func(stage overlaystage.Overlay) error {
				return stage.Set(workbookPath, workbook)
			})
			if err == nil {
				t.Fatalf("expected postflight error")
			}

			if tc.mode == BestEffort {
				alerts := doc.AlertsByCode("POSTFLIGHT_XLSX_CELL_TYPE_MISMATCH")
				if len(alerts) != 1 {
					t.Fatalf("expected POSTFLIGHT_XLSX_CELL_TYPE_MISMATCH alert, got %d", len(alerts))
				}
			}

			if err := doc.SaveFile(outputPath); err != nil {
				t.Fatalf("SaveFile: %v", err)
			}
			outWorkbook := readZipEntry(t, outputPath, workbookPath)
			if !bytes.Equal(outWorkbook, workbook) {
				t.Fatalf("workbook changed after discard")
			}
		})
	}
}

func buildXLSXWithSharedStringCell(t *testing.T) []byte {
	t.Helper()

	parts := map[string][]byte{
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

	return writeZipBytes(t, parts)
}
