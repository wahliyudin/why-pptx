package pptx

import (
	"path/filepath"
	"testing"

	"why-pptx/internal/testutil/pptxassert"
)

func TestDeterminismDoubleRun(t *testing.T) {
	input := fixturePath("mix_write_secondary_axis_valid.pptx")
	data := map[string][]string{
		"categories": {"CatA", "CatB"},
		"values:0":   {"101", "202"},
		"values:1":   {"303", "404"},
	}

	first := buildSnapshotAfterApply(t, input, data)
	second := buildSnapshotAfterApply(t, input, data)

	pptxassert.AssertSnapshotEqual(t, first, second)
}

func buildSnapshotAfterApply(t *testing.T, input string, data map[string][]string) pptxassert.Snapshot {
	t.Helper()

	doc, err := OpenFile(input)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}
	if err := doc.ApplyChartDataByPath("ppt/charts/chart1.xml", data); err != nil {
		t.Fatalf("ApplyChartDataByPath: %v", err)
	}

	output := filepath.Join(t.TempDir(), "output.pptx")
	if err := doc.SaveFile(output); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	snap, err := pptxassert.BuildSnapshot(output)
	if err != nil {
		t.Fatalf("BuildSnapshot: %v", err)
	}
	return snap
}
