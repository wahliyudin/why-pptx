package pptx

import "testing"

func TestChartJSExporterMissingNumericEmpty(t *testing.T) {
	exporter := ChartJSExporter{MissingNumericPolicy: MissingNumericEmpty}
	input := ExtractedChartData{
		Type:   "bar",
		Labels: []string{"A", "B"},
		Series: []ExtractedSeries{{
			Index: 0,
			Name:  "Series 1",
			Data:  []string{"1", ""},
		}},
	}

	payload, err := exporter.Export(input)
	if err != nil {
		t.Fatalf("Export: %v", err)
	}
	if payload.Format != ExportChartJS {
		t.Fatalf("unexpected format %q", payload.Format)
	}

	datasets, ok := payload.Data["datasets"].([]map[string]any)
	if !ok || len(datasets) != 1 {
		t.Fatalf("unexpected datasets: %#v", payload.Data["datasets"])
	}
	values, ok := datasets[0]["data"].([]any)
	if !ok || len(values) != 2 {
		t.Fatalf("unexpected data values: %#v", datasets[0]["data"])
	}
	if val, ok := values[0].(float64); !ok || val != 1 {
		t.Fatalf("unexpected first value: %#v", values[0])
	}
	if values[1] != nil {
		t.Fatalf("expected nil for empty numeric, got %#v", values[1])
	}
}

func TestChartJSExporterMissingNumericZero(t *testing.T) {
	exporter := ChartJSExporter{MissingNumericPolicy: MissingNumericZero}
	input := ExtractedChartData{
		Type:   "line",
		Labels: []string{"A", "B", "C"},
		Series: []ExtractedSeries{{
			Index: 0,
			Name:  "Series 1",
			Data:  []string{"", "2", "bad"},
		}},
	}

	payload, err := exporter.Export(input)
	if err != nil {
		t.Fatalf("Export: %v", err)
	}
	datasets := payload.Data["datasets"].([]map[string]any)
	values := datasets[0]["data"].([]any)
	if val, ok := values[0].(float64); !ok || val != 0 {
		t.Fatalf("unexpected value[0]: %#v", values[0])
	}
	if val, ok := values[1].(float64); !ok || val != 2 {
		t.Fatalf("unexpected value[1]: %#v", values[1])
	}
	if val, ok := values[2].(float64); !ok || val != 0 {
		t.Fatalf("unexpected value[2]: %#v", values[2])
	}
}

func TestChartJSExporterInvalidNumeric(t *testing.T) {
	exporter := ChartJSExporter{MissingNumericPolicy: MissingNumericEmpty}
	input := ExtractedChartData{
		Type:   "bar",
		Labels: []string{"A"},
		Series: []ExtractedSeries{{
			Index: 0,
			Name:  "Series 1",
			Data:  []string{"bad"},
		}},
	}

	if _, err := exporter.Export(input); err == nil {
		t.Fatalf("expected export error")
	}
}

func TestChartJSExporterPie(t *testing.T) {
	exporter := ChartJSExporter{MissingNumericPolicy: MissingNumericEmpty}
	input := ExtractedChartData{
		Type:   "pie",
		Labels: []string{"A", "B"},
		Series: []ExtractedSeries{{
			Index: 0,
			Name:  "Series 1",
			Data:  []string{"1", "2"},
		}},
	}

	payload, err := exporter.Export(input)
	if err != nil {
		t.Fatalf("Export: %v", err)
	}
	if payload.Data["type"] != "pie" {
		t.Fatalf("unexpected type: %#v", payload.Data["type"])
	}
	datasets := payload.Data["datasets"].([]map[string]any)
	if len(datasets) != 1 {
		t.Fatalf("expected 1 dataset, got %d", len(datasets))
	}
}

func TestChartJSExporterArea(t *testing.T) {
	exporter := ChartJSExporter{MissingNumericPolicy: MissingNumericEmpty}
	input := ExtractedChartData{
		Type:   "area",
		Labels: []string{"A"},
		Series: []ExtractedSeries{{
			Index: 0,
			Name:  "Series 1",
			Data:  []string{"1"},
		}},
	}

	payload, err := exporter.Export(input)
	if err != nil {
		t.Fatalf("Export: %v", err)
	}
	if payload.Data["type"] != "line" {
		t.Fatalf("expected line type for area export, got %#v", payload.Data["type"])
	}
	datasets := payload.Data["datasets"].([]map[string]any)
	if len(datasets) != 1 {
		t.Fatalf("expected 1 dataset, got %d", len(datasets))
	}
	if fill, ok := datasets[0]["fill"].(bool); !ok || !fill {
		t.Fatalf("expected fill=true for area dataset")
	}
}
