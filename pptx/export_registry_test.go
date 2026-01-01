package pptx

import (
	"reflect"
	"testing"
)

type dummyExporter struct {
	format ExportFormat
}

func (d dummyExporter) Format() ExportFormat { return d.format }

func (d dummyExporter) Export(in ExtractedChartData) (ExportedPayload, error) {
	return ExportedPayload{Format: d.format}, nil
}

func TestExporterRegistryRegisterGet(t *testing.T) {
	reg := NewExporterRegistry()
	if err := reg.Register(ChartJSExporter{}); err != nil {
		t.Fatalf("Register: %v", err)
	}
	got, ok := reg.Get(ExportChartJS)
	if !ok {
		t.Fatalf("expected exporter")
	}
	if got.Format() != ExportChartJS {
		t.Fatalf("unexpected format %q", got.Format())
	}
}

func TestExporterRegistryDuplicate(t *testing.T) {
	reg := NewExporterRegistry()
	if err := reg.Register(ChartJSExporter{}); err != nil {
		t.Fatalf("Register: %v", err)
	}
	if err := reg.Register(ChartJSExporter{}); err == nil {
		t.Fatalf("expected duplicate registration error")
	}
}

func TestExporterRegistryFormatsSorted(t *testing.T) {
	reg := NewExporterRegistry()
	if err := reg.Register(dummyExporter{format: ExportChartJS}); err != nil {
		t.Fatalf("Register: %v", err)
	}
	if err := reg.Register(dummyExporter{format: "d3"}); err != nil {
		t.Fatalf("Register: %v", err)
	}

	formats := reg.Formats()
	want := []ExportFormat{ExportChartJS, "d3"}
	if !reflect.DeepEqual(formats, want) {
		t.Fatalf("formats mismatch: got %v want %v", formats, want)
	}
}

func TestDefaultExporterRegistryHasChartJS(t *testing.T) {
	reg := DefaultExporterRegistry()
	if _, ok := reg.Get(ExportChartJS); !ok {
		t.Fatalf("expected chartjs exporter in default registry")
	}
}

func TestExportChartByPathFormatMatchesExporter(t *testing.T) {
	doc, err := OpenFile(fixturePath("bar_simple_embedded.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	byFormat, err := doc.ExportChartByPathFormat("ppt/charts/chart1.xml", ExportChartJS)
	if err != nil {
		t.Fatalf("ExportChartByPathFormat: %v", err)
	}

	explicit, err := doc.ExportChartByPath("ppt/charts/chart1.xml", ChartJSExporter{
		MissingNumericPolicy: doc.opts.Workbook.MissingNumericPolicy,
	})
	if err != nil {
		t.Fatalf("ExportChartByPath: %v", err)
	}

	if !reflect.DeepEqual(byFormat, explicit) {
		t.Fatalf("payload mismatch")
	}
}

func TestExportChartByPathFormatUnknown(t *testing.T) {
	opts := DefaultOptions()
	opts.Mode = BestEffort
	doc, err := OpenFile(fixturePath("bar_simple_embedded.pptx"), WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if _, err := doc.ExportChartByPathFormat("ppt/charts/chart1.xml", "unknown"); err == nil {
		t.Fatalf("expected export error")
	}

	alerts := doc.AlertsByCode("EXPORT_FORMAT_UNSUPPORTED")
	if len(alerts) != 1 {
		t.Fatalf("expected EXPORT_FORMAT_UNSUPPORTED alert, got %d", len(alerts))
	}
}
