package pptx

import (
	"errors"
	"flag"
	"os"
	"path/filepath"
	"testing"

	"why-pptx/internal/testutil/pptxassert"
)

var updateGolden = flag.Bool("update-golden", false, "update golden snapshots")

func TestGoldenSnapshots(t *testing.T) {
	cases := []struct {
		name  string
		input string
		build func(outputPath string) error
	}{
		{
			name:  "bar_simple_update",
			input: fixturePath("bar_simple_embedded.pptx"),
			build: func(outputPath string) error {
				doc, err := OpenFile(fixturePath("bar_simple_embedded.pptx"))
				if err != nil {
					return err
				}
				data := map[string][]string{
					"categories": {"New1", "New2"},
					"values:0":   {"100", "200"},
				}
				if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
					return err
				}
				return doc.SaveFile(outputPath)
			},
		},
		{
			name:  "mix_write_bar_line_update",
			input: fixturePath("mix_write_bar_line_valid.pptx"),
			build: func(outputPath string) error {
				doc, err := OpenFile(fixturePath("mix_write_bar_line_valid.pptx"))
				if err != nil {
					return err
				}
				data := map[string][]string{
					"categories": {"CatA", "CatB"},
					"values:0":   {"11", "22"},
					"values:1":   {"33", "44"},
				}
				if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
					return err
				}
				return doc.SaveFile(outputPath)
			},
		},
		{
			name:  "mix_write_secondary_axis_update",
			input: fixturePath("mix_write_secondary_axis_valid.pptx"),
			build: func(outputPath string) error {
				doc, err := OpenFile(fixturePath("mix_write_secondary_axis_valid.pptx"))
				if err != nil {
					return err
				}
				data := map[string][]string{
					"categories": {"CatA", "CatB"},
					"values:0":   {"101", "202"},
					"values:1":   {"303", "404"},
				}
				if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
					return err
				}
				return doc.SaveFile(outputPath)
			},
		},
		{
			name:  "mix_write_secondary_axis_variantA",
			input: fixturePath("mix_write_secondary_axis_valid_variantA.pptx"),
		},
		{
			name:  "mix_write_secondary_axis_variantB",
			input: fixturePath("mix_write_secondary_axis_valid_variantB.pptx"),
		},
		{
			name:  "line_chart_cached_values_missing",
			input: fixturePath("line_chart_cached_values_missing.pptx"),
		},
		{
			name:  "workbook_inlineStr_edgecases",
			input: fixturePath("workbook_inlineStr_edgecases.pptx"),
		},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			outputPath := tc.input
			if tc.build != nil {
				outputPath = filepath.Join(t.TempDir(), "output.pptx")
				if err := tc.build(outputPath); err != nil {
					t.Fatalf("build %s: %v", tc.name, err)
				}
			}

			snap, err := pptxassert.BuildSnapshot(outputPath)
			if err != nil {
				t.Fatalf("BuildSnapshot: %v", err)
			}

			goldenPath := filepath.Join("..", "testdata", "golden", tc.name+".json")
			if *updateGolden {
				if err := pptxassert.WriteSnapshot(goldenPath, snap); err != nil {
					t.Fatalf("WriteSnapshot: %v", err)
				}
				return
			}

			want, err := pptxassert.LoadSnapshot(goldenPath)
			if err != nil {
				if errors.Is(err, os.ErrNotExist) {
					t.Fatalf("golden snapshot missing: %s (run tests with -update-golden)", goldenPath)
				}
				t.Fatalf("LoadSnapshot: %v", err)
			}
			pptxassert.AssertSnapshotEqual(t, snap, want)
		})
	}
}
