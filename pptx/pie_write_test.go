package pptx

import (
	"bytes"
	"path/filepath"
	"testing"

	"why-pptx/internal/overlaystage"
	"why-pptx/internal/postflight"
	"why-pptx/internal/testutil/pptxassert"
)

func TestPieApplyChartDataStrict(t *testing.T) {
	input := fixturePath("pie_edit_valid.pptx")
	output := filepath.Join(t.TempDir(), "output.pptx")

	doc, err := OpenFile(input)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2", "New3"},
		"values:0":   {"11", "22", "33"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
		t.Fatalf("ApplyChartDataByPath: %v", err)
	}
	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	pptxassert.AssertSameEntrySet(t, input, output)

	chartXML, err := pptxassert.ReadEntry(output, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ReadEntry chart: %v", err)
	}
	snap, err := pptxassert.ExtractChartCacheSnapshot(chartXML)
	if err != nil {
		t.Fatalf("ExtractChartCacheSnapshot: %v", err)
	}
	pptxassert.AssertCacheMatchesExpected(t, snap, pptxassert.ExpectedCache{
		Series: []pptxassert.ExpectedCacheSeries{
			{Kind: "strCache", SeriesIndex: 0, Values: []string{"New1", "New2", "New3"}},
			{Kind: "numCache", SeriesIndex: 0, Values: []string{"11", "22", "33"}},
		},
	})

	workbook, err := pptxassert.ReadEntry(output, "ppt/embeddings/embeddedWorkbook1.xlsx")
	if err != nil {
		t.Fatalf("ReadEntry workbook: %v", err)
	}
	cells, err := pptxassert.ExtractWorkbookCellSnapshot(workbook, "Sheet1", []string{"A2", "A3", "A4", "B2", "B3", "B4"})
	if err != nil {
		t.Fatalf("ExtractWorkbookCellSnapshot: %v", err)
	}
	if cells["A2"] != "New1" || cells["A3"] != "New2" || cells["A4"] != "New3" {
		t.Fatalf("unexpected category cells: %v", cells)
	}
	if cells["B2"] != "11" || cells["B3"] != "22" || cells["B4"] != "33" {
		t.Fatalf("unexpected value cells: %v", cells)
	}
}

func TestPieApplyChartDataMultipleSeriesStrict(t *testing.T) {
	doc, err := OpenFile(fixturePath("pie_edit_multiple_series.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"1", "2"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}
}

func TestPieApplyChartDataMultipleSeriesBestEffort(t *testing.T) {
	input := fixturePath("pie_edit_multiple_series.pptx")
	output := filepath.Join(t.TempDir(), "output.pptx")

	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(input, WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"1", "2"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}

	alerts := doc.AlertsByCode("WRITE_PIE_MULTIPLE_SERIES_UNSUPPORTED")
	if len(alerts) != 1 {
		t.Fatalf("expected WRITE_PIE_MULTIPLE_SERIES_UNSUPPORTED alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}
	pptxassert.AssertSameEntrySet(t, input, output)

	beforeChart, err := pptxassert.ReadEntry(input, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ReadEntry chart before: %v", err)
	}
	afterChart, err := pptxassert.ReadEntry(output, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ReadEntry chart after: %v", err)
	}
	if !bytes.Equal(beforeChart, afterChart) {
		t.Fatalf("chart xml changed after skip")
	}
}

func TestPieApplyChartDataLinkedWorkbookBestEffort(t *testing.T) {
	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(fixturePath("pie_edit_linked_workbook.pptx"), WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"1", "2"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}

	alerts := doc.AlertsByCode("CHART_LINKED_WORKBOOK")
	if len(alerts) != 1 {
		t.Fatalf("expected CHART_LINKED_WORKBOOK alert, got %d", len(alerts))
	}
}

func TestPiePostflightInvalidCacheStrict(t *testing.T) {
	input := fixturePath("pie_edit_cache_invalid.pptx")
	output := filepath.Join(t.TempDir(), "output.pptx")

	doc, err := OpenFile(input)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	chartPath := "ppt/charts/chart1.xml"
	badChart, err := pptxassert.ReadEntry(input, chartPath)
	if err != nil {
		t.Fatalf("ReadEntry chart: %v", err)
	}

	ctx := postflight.ValidateContext{
		ChartPath:        chartPath,
		Mode:             postflight.ModeStrict,
		CacheSyncEnabled: true,
	}
	err = doc.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		return stage.Set(chartPath, badChart)
	})
	if err == nil {
		t.Fatalf("expected postflight error")
	}

	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}
	pptxassert.AssertSameEntrySet(t, input, output)
}
