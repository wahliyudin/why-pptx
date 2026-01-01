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
	if in.Type != "bar" && in.Type != "line" {
		return ExportedPayload{}, fmt.Errorf("unsupported chart type %q", in.Type)
	}

	series := make([]ExtractedSeries, len(in.Series))
	copy(series, in.Series)
	sort.Slice(series, func(i, j int) bool {
		return series[i].Index < series[j].Index
	})

	datasets := make([]map[string]any, 0, len(series))
	for _, s := range series {
		values := make([]any, len(s.Data))
		for i, raw := range s.Data {
			val, err := chartJSValue(raw, e.MissingNumericPolicy)
			if err != nil {
				return ExportedPayload{}, fmt.Errorf("series %d value %q: %w", s.Index, raw, err)
			}
			values[i] = val
		}
		datasets = append(datasets, map[string]any{
			"label": s.Name,
			"data":  values,
		})
	}

	labels := append([]string(nil), in.Labels...)
	return ExportedPayload{
		Format: ExportChartJS,
		Data: map[string]any{
			"type":     in.Type,
			"labels":   labels,
			"datasets": datasets,
		},
	}, nil
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
