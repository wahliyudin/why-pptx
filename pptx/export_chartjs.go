package pptx

import (
	"fmt"
	"sort"
	"strconv"
	"strings"
)

// ChartJSExporter builds a minimal Chart.js payload from extracted chart data.
type ChartJSExporter struct {
	MissingNumericPolicy MissingNumericPolicy
}

func (e ChartJSExporter) Format() ExportFormat {
	return ExportChartJS
}

func (e ChartJSExporter) Export(in ExtractedChartData) (ExportedPayload, error) {
	switch in.Type {
	case "bar", "line", "pie", "area", "mixed":
	default:
		return ExportedPayload{}, fmt.Errorf("unsupported chart type %q", in.Type)
	}

	series := make([]ExtractedSeries, len(in.Series))
	copy(series, in.Series)
	sort.Slice(series, func(i, j int) bool {
		return series[i].Index < series[j].Index
	})

	if in.Type == "pie" {
		if len(series) != 1 {
			return ExportedPayload{}, fmt.Errorf("pie chart requires a single series")
		}
		values, err := chartJSValues(series[0].Index, series[0].Data, e.MissingNumericPolicy)
		if err != nil {
			return ExportedPayload{}, err
		}
		labels := append([]string(nil), in.Labels...)
		return ExportedPayload{
			Format: ExportChartJS,
			Data: map[string]any{
				"type":   "pie",
				"labels": labels,
				"datasets": []map[string]any{{
					"label": series[0].Name,
					"data":  values,
				}},
			},
		}, nil
	}

	if in.Type == "mixed" {
		if len(series) == 0 {
			return ExportedPayload{}, fmt.Errorf("mixed chart requires at least one series")
		}
		chartType := "line"
		for _, s := range series {
			if s.PlotType == "bar" {
				chartType = "bar"
				break
			}
		}

		datasets := make([]map[string]any, 0, len(series))
		for _, s := range series {
			if s.PlotType != "bar" && s.PlotType != "line" {
				return ExportedPayload{}, fmt.Errorf("mixed chart series %d has unsupported plot type %q", s.Index, s.PlotType)
			}
			values, err := chartJSValues(s.Index, s.Data, e.MissingNumericPolicy)
			if err != nil {
				return ExportedPayload{}, err
			}
			dataset := map[string]any{
				"label": s.Name,
				"data":  values,
				"type":  s.PlotType,
			}
			datasets = append(datasets, dataset)
		}

		labels := append([]string(nil), in.Labels...)
		return ExportedPayload{
			Format: ExportChartJS,
			Data: map[string]any{
				"type":     chartType,
				"labels":   labels,
				"datasets": datasets,
			},
		}, nil
	}

	chartType := in.Type
	fill := false
	if in.Type == "area" {
		chartType = "line"
		fill = true
	}

	datasets := make([]map[string]any, 0, len(series))
	for _, s := range series {
		values, err := chartJSValues(s.Index, s.Data, e.MissingNumericPolicy)
		if err != nil {
			return ExportedPayload{}, err
		}
		dataset := map[string]any{
			"label": s.Name,
			"data":  values,
		}
		if fill {
			dataset["fill"] = true
		}
		datasets = append(datasets, dataset)
	}

	labels := append([]string(nil), in.Labels...)
	return ExportedPayload{
		Format: ExportChartJS,
		Data: map[string]any{
			"type":     chartType,
			"labels":   labels,
			"datasets": datasets,
		},
	}, nil
}

func chartJSValues(seriesIndex int, values []string, policy MissingNumericPolicy) ([]any, error) {
	out := make([]any, len(values))
	for i, raw := range values {
		val, err := chartJSValue(raw, policy)
		if err != nil {
			return nil, fmt.Errorf("series %d value %q: %w", seriesIndex, raw, err)
		}
		out[i] = val
	}
	return out, nil
}

func chartJSValue(raw string, policy MissingNumericPolicy) (any, error) {
	trimmed := strings.TrimSpace(raw)
	if trimmed == "" {
		if policy == MissingNumericZero {
			return float64(0), nil
		}
		return nil, nil
	}

	value, err := strconv.ParseFloat(trimmed, 64)
	if err != nil {
		if policy == MissingNumericZero {
			return float64(0), nil
		}
		return nil, fmt.Errorf("invalid number %q", raw)
	}
	return value, nil
}
