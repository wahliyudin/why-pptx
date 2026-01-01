package pptx

import (
	"bytes"
	"path/filepath"
	"testing"

	"why-pptx/internal/overlaystage"
	"why-pptx/internal/postflight"
	"why-pptx/internal/testutil/pptxassert"
)

func TestStrictBarSimpleUpdate_PreservesStructure(t *testing.T) {
	input := fixturePath("bar_simple_embedded.pptx")
	output := filepath.Join(t.TempDir(), "output.pptx")

	doc, err := OpenFile(input)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"100", "200"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
		t.Fatalf("ApplyChartDataByPath: %v", err)
	}
	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	pptxassert.AssertSameEntrySet(t, input, output)
	pptxassert.AssertRelTargetsExist(t, output, []string{
		"ppt/slides/_rels/slide1.xml.rels",
		"ppt/charts/_rels/chart1.xml.rels",
	})

	workbook, err := pptxassert.ReadEntry(output, "ppt/embeddings/embeddedWorkbook1.xlsx")
	if err != nil {
		t.Fatalf("ReadEntry workbook: %v", err)
	}
	pptxassert.AssertNoSharedStringsPart(t, workbook)
	pptxassert.AssertNoSharedStringCells(t, workbook)

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
			{Kind: "numCache", SeriesIndex: 0, Values: []string{"100", "200"}},
		},
	})

	cells, err := pptxassert.ExtractWorkbookCellSnapshot(workbook, "Sheet1", []string{"A2", "A3", "B2", "B3"})
	if err != nil {
		t.Fatalf("ExtractWorkbookCellSnapshot: %v", err)
	}
	if cells["A2"] != "New1" || cells["A3"] != "New2" {
		t.Fatalf("unexpected category cells: %v", cells)
	}
	if cells["B2"] != "100" || cells["B3"] != "200" {
		t.Fatalf("unexpected value cells: %v", cells)
	}
}

func TestBestEffortLinkedWorkbook_SkipsWithAlert_NoChanges(t *testing.T) {
	input := fixturePath("linked_workbook_chart.pptx")
	output := filepath.Join(t.TempDir(), "output.pptx")

	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(input, WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	charts, err := doc.ListCharts()
	if err != nil {
		t.Fatalf("ListCharts: %v", err)
	}
	if len(charts) != 0 {
		t.Fatalf("expected no embedded charts, got %d", len(charts))
	}

	alerts := doc.AlertsByCode("CHART_LINKED_WORKBOOK")
	if len(alerts) != 1 {
		t.Fatalf("expected CHART_LINKED_WORKBOOK alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}
	pptxassert.AssertSameEntrySet(t, input, output)
}

func TestStrictMalformedCache_Abort(t *testing.T) {
	input := fixturePath("malformed_chart_cache.pptx")
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

	outChart, err := pptxassert.ReadEntry(output, chartPath)
	if err != nil {
		t.Fatalf("ReadEntry output chart: %v", err)
	}
	if !bytes.Equal(outChart, badChart) {
		t.Fatalf("chart xml changed after abort")
	}
}

func TestBestEffortMalformedCache_Skip_AndContinue(t *testing.T) {
	input := fixturePath("shared_workbook_two_charts.pptx")
	output := filepath.Join(t.TempDir(), "output.pptx")

	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(input, WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	data := map[string][]string{
		"categories": {"New1", "New2"},
		"values:0":   {"111", "222"},
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
		t.Fatalf("ApplyChartDataByPath chart1: %v", err)
	}

	badChart, err := pptxassert.ReadEntry(fixturePath("malformed_chart_cache.pptx"), "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ReadEntry malformed chart: %v", err)
	}

	ctx := postflight.ValidateContext{
		ChartPath:        "ppt/charts/chart2.xml",
		Mode:             postflight.ModeBestEffort,
		CacheSyncEnabled: true,
	}
	err = doc.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		return stage.Set("ppt/charts/chart2.xml", badChart)
	})
	if err == nil {
		t.Fatalf("expected postflight error for chart2")
	}

	alerts := doc.AlertsByCode("POSTFLIGHT_CHART_CACHE_INVALID")
	if len(alerts) != 1 {
		t.Fatalf("expected POSTFLIGHT_CHART_CACHE_INVALID alert, got %d", len(alerts))
	}

	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}
	pptxassert.AssertSameEntrySet(t, input, output)

	chart1XML, err := pptxassert.ReadEntry(output, "ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("ReadEntry chart1: %v", err)
	}
	chart1Snap, err := pptxassert.ExtractChartCacheSnapshot(chart1XML)
	if err != nil {
		t.Fatalf("ExtractChartCacheSnapshot: %v", err)
	}
	pptxassert.AssertCacheMatchesExpected(t, chart1Snap, pptxassert.ExpectedCache{
		Series: []pptxassert.ExpectedCacheSeries{
			{Kind: "strCache", SeriesIndex: 0, Values: []string{"New1", "New2"}},
			{Kind: "numCache", SeriesIndex: 0, Values: []string{"111", "222"}},
		},
	})

	chart2Before, err := pptxassert.ReadEntry(input, "ppt/charts/chart2.xml")
	if err != nil {
		t.Fatalf("ReadEntry chart2 before: %v", err)
	}
	chart2After, err := pptxassert.ReadEntry(output, "ppt/charts/chart2.xml")
	if err != nil {
		t.Fatalf("ReadEntry chart2 after: %v", err)
	}
	if !bytes.Equal(chart2Before, chart2After) {
		t.Fatalf("chart2 xml changed after skip")
	}

	workbook, err := pptxassert.ReadEntry(output, "ppt/embeddings/embeddedWorkbook1.xlsx")
	if err != nil {
		t.Fatalf("ReadEntry workbook: %v", err)
	}
	cells, err := pptxassert.ExtractWorkbookCellSnapshot(workbook, "Sheet1", []string{"A2", "A3", "B2", "B3"})
	if err != nil {
		t.Fatalf("ExtractWorkbookCellSnapshot: %v", err)
	}
	if cells["A2"] != "New1" || cells["A3"] != "New2" {
		t.Fatalf("unexpected category cells: %v", cells)
	}
	if cells["B2"] != "111" || cells["B3"] != "222" {
		t.Fatalf("unexpected value cells: %v", cells)
	}
}

func TestSharedStringsFatal(t *testing.T) {
	cases := []struct {
		name string
		mode ErrorMode
	}{
		{name: "best-effort", mode: BestEffort},
		{name: "strict", mode: Strict},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			input := fixturePath("xlsx_sharedStrings_present.pptx")
			output := filepath.Join(t.TempDir(), "output.pptx")

			opts := DefaultOptions()
			opts.Mode = tc.mode
			doc, err := OpenFile(input, WithOptions(opts))
			if err != nil {
				t.Fatalf("OpenFile: %v", err)
			}

			data := map[string][]string{
				"categories": {"New1", "New2"},
				"values:0":   {"10", "20"},
			}
			err = doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data)
			if err == nil {
				t.Fatalf("expected ApplyChartDataByPath error")
			}

			alerts := doc.AlertsByCode("POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED")
			if len(alerts) != 1 {
				t.Fatalf("expected POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED alert, got %d", len(alerts))
			}

			if err := doc.SaveFile(output); err != nil {
				t.Fatalf("SaveFile: %v", err)
			}
			pptxassert.AssertSameEntrySet(t, input, output)
		})
	}
}

func fixturePath(name string) string {
	return filepath.Join("..", "testdata", "pptx", name)
}
