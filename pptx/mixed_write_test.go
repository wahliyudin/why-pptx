package pptx

import (
	"bytes"
	"path/filepath"
	"testing"

	"why-pptx/internal/overlaystage"
	"why-pptx/internal/postflight"
	"why-pptx/internal/testutil/pptxassert"
)

func TestMixedApplyChartDataStrict(t *testing.T) {
	input := fixturePath("mix_write_bar_line_valid.pptx")
	output := filepath.Join(t.TempDir(), "output.pptx")

	doc, err := OpenFile(input)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"11", "22"},
		"values:1":   {"33", "44"},
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
			{Kind: "strCache", SeriesIndex: 0, Values: []string{"New1", "New2"}},
			{Kind: "numCache", SeriesIndex: 0, Values: []string{"11", "22"}},
			{Kind: "strCache", SeriesIndex: 1, Values: []string{"New1", "New2"}},
			{Kind: "numCache", SeriesIndex: 1, Values: []string{"33", "44"}},
		},
	})

	workbook, err := pptxassert.ReadEntry(output, "ppt/embeddings/embeddedWorkbook1.xlsx")
	if err != nil {
		t.Fatalf("ReadEntry workbook: %v", err)
	}
	cells, err := pptxassert.ExtractWorkbookCellSnapshot(workbook, "Sheet1", []string{"A2", "A3", "B2", "B3", "C2", "C3"})
	if err != nil {
		t.Fatalf("ExtractWorkbookCellSnapshot: %v", err)
	}
	if cells["A2"] != "New1" || cells["A3"] != "New2" {
		t.Fatalf("unexpected category cells: %v", cells)
	}
	if cells["B2"] != "11" || cells["B3"] != "22" {
		t.Fatalf("unexpected bar series cells: %v", cells)
	}
	if cells["C2"] != "33" || cells["C3"] != "44" {
		t.Fatalf("unexpected line series cells: %v", cells)
	}
}

func TestMixedApplyChartDataSecondaryAxisStrict(t *testing.T) {
	doc, err := OpenFile(fixturePath("mix_write_secondary_axis.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"1", "2"},
		"values:1":   {"3", "4"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}
}

func TestMixedApplyChartDataSecondaryAxisBestEffort(t *testing.T) {
	input := fixturePath("mix_write_secondary_axis.pptx")
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
		"values:1":   {"3", "4"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}

	alerts := doc.AlertsByCode("WRITE_MIX_SECONDARY_AXIS_UNSUPPORTED")
	if len(alerts) != 1 {
		t.Fatalf("expected WRITE_MIX_SECONDARY_AXIS_UNSUPPORTED alert, got %d", len(alerts))
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

func TestMixedApplyChartDataUnsupportedPlotStrict(t *testing.T) {
	doc, err := OpenFile(fixturePath("mix_unsupported_variant.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"1", "2"},
		"values:1":   {"3", "4"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}
}

func TestMixedApplyChartDataUnsupportedPlotBestEffort(t *testing.T) {
	input := fixturePath("mix_unsupported_variant.pptx")
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
		"values:1":   {"3", "4"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}

	alerts := doc.AlertsByCode("WRITE_MIX_UNSUPPORTED_SHAPE")
	if len(alerts) != 1 {
		t.Fatalf("expected WRITE_MIX_UNSUPPORTED_SHAPE alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}
	pptxassert.AssertSameEntrySet(t, input, output)
}

func TestMixedApplyChartDataMismatchedCategoriesStrict(t *testing.T) {
	doc, err := OpenFile(fixturePath("mix_write_mismatched_categories.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"1", "2"},
		"values:1":   {"3", "4"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}
}

func TestMixedApplyChartDataMismatchedCategoriesBestEffort(t *testing.T) {
	input := fixturePath("mix_write_mismatched_categories.pptx")
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
		"values:1":   {"3", "4"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err == nil {
		t.Fatalf("expected ApplyChartDataByPath error")
	}

	alerts := doc.AlertsByCode("CHART_DEPENDENCIES_PARSE_FAILED")
	if len(alerts) != 1 {
		t.Fatalf("expected CHART_DEPENDENCIES_PARSE_FAILED alert, got %d", len(alerts))
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

func TestMixedApplyChartDataLengthMismatchBestEffort(t *testing.T) {
	input := fixturePath("mix_write_bar_line_valid.pptx")
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
		"values:1":   {"3", "4", "5"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
		t.Fatalf("expected ApplyChartDataByPath to skip without error, got %v", err)
	}

	alerts := doc.AlertsByCode("CHART_DATA_LENGTH_MISMATCH")
	if len(alerts) != 1 {
		t.Fatalf("expected CHART_DATA_LENGTH_MISMATCH alert, got %d", len(alerts))
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
		t.Fatalf("chart xml changed after length mismatch skip")
	}
}

func TestMixedPostflightInvalidCacheStrict(t *testing.T) {
	input := fixturePath("mix_write_cache_invalid.pptx")
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
