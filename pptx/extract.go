package pptx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strings"

	"why-pptx/internal/chartdiscover"
	"why-pptx/internal/chartxml"
	"why-pptx/internal/xlref"
	"why-pptx/internal/xlsxembed"
)

type ExtractedChartData struct {
	Type   string            `json:"type"`
	Labels []string          `json:"labels"`
	Series []ExtractedSeries `json:"series"`
	Meta   ExtractMeta       `json:"meta"`
}

type ExtractedSeries struct {
	Index int      `json:"index"`
	Name  string   `json:"name"`
	Data  []string `json:"data"`
	// PlotType is set for mixed charts (e.g., "bar" or "line").
	PlotType string `json:"plotType,omitempty"`
	// Axis is set for mixed charts when a secondary axis is detected.
	Axis string `json:"axis,omitempty"`
}

type ExtractMeta struct {
	ChartPath    string `json:"chartPath"`
	SlidePath    string `json:"slidePath"`
	WorkbookPath string `json:"workbookPath"`
	Sheet        string `json:"sheet,omitempty"`
}

type ExportFormat string

const (
	ExportChartJS ExportFormat = "chartjs"
	ExportD3      ExportFormat = "d3"
)

type ExportedPayload struct {
	Format ExportFormat   `json:"format"`
	Data   map[string]any `json:"data"`
}

type Exporter interface {
	Format() ExportFormat
	Export(in ExtractedChartData) (ExportedPayload, error)
}

type mixedSeriesRanges struct {
	series     chartxml.MixedSeries
	categories *ChartRange
	values     *ChartRange
	name       *ChartRange
}

func (d *Document) ExtractChartDataByPath(chartPath string) (ExtractedChartData, error) {
	if d == nil || d.pkg == nil {
		return ExtractedChartData{}, fmt.Errorf("document not initialized")
	}
	if chartPath == "" {
		return ExtractedChartData{}, fmt.Errorf("chart path is required")
	}

	embedded, skipped, err := chartdiscover.DiscoverEmbeddedCharts(d.pkg)
	if err != nil {
		return ExtractedChartData{}, err
	}

	for _, skip := range skipped {
		if skip.ChartPath == chartPath {
			return ExtractedChartData{}, d.handleExtractError(extractIssue{
				code:    mapSkipReasonCode(skip),
				message: extractMessageForCode(mapSkipReasonCode(skip)),
				err:     fmt.Errorf("chart %q is not eligible for extraction", chartPath),
				context: extractSkipContext(skip),
			})
		}
	}

	for _, item := range embedded {
		if item.ChartPath == chartPath {
			data, err := d.extractChartData(item)
			if err != nil {
				return ExtractedChartData{}, err
			}
			return data, nil
		}
	}

	return ExtractedChartData{}, fmt.Errorf("chart not found")
}

func (d *Document) ExtractChartData(chartIndex int) (ExtractedChartData, error) {
	if d == nil || d.pkg == nil {
		return ExtractedChartData{}, fmt.Errorf("document not initialized")
	}
	if chartIndex < 0 {
		return ExtractedChartData{}, fmt.Errorf("chart index out of range")
	}

	charts, err := d.DiscoverEmbeddedCharts()
	if err != nil {
		return ExtractedChartData{}, err
	}
	if chartIndex >= len(charts) {
		return ExtractedChartData{}, fmt.Errorf("chart index out of range")
	}

	return d.extractChartData(chartdiscover.EmbeddedChart{
		SlidePath:    charts[chartIndex].SlidePath,
		ChartPath:    charts[chartIndex].ChartPath,
		WorkbookPath: charts[chartIndex].WorkbookPath,
	})
}

func (d *Document) ExtractAllCharts() ([]ExtractedChartData, error) {
	if d == nil || d.pkg == nil {
		return nil, fmt.Errorf("document not initialized")
	}

	embedded, skipped, err := chartdiscover.DiscoverEmbeddedCharts(d.pkg)
	if err != nil {
		return nil, err
	}

	out := make([]ExtractedChartData, 0, len(embedded))

	for _, skip := range skipped {
		err := d.handleExtractError(extractIssue{
			code:    mapSkipReasonCode(skip),
			message: extractMessageForCode(mapSkipReasonCode(skip)),
			err:     fmt.Errorf("chart %q is not eligible for extraction", skip.ChartPath),
			context: extractSkipContext(skip),
		})
		if err != nil && d.opts.Mode == Strict {
			return nil, err
		}
	}

	for _, chart := range embedded {
		data, err := d.extractChartData(chart)
		if err != nil {
			if d.opts.Mode == BestEffort {
				continue
			}
			return nil, err
		}
		out = append(out, data)
	}

	if len(out) == 0 {
		return []ExtractedChartData{}, nil
	}
	return out, nil
}

func (d *Document) ExportChartByPath(chartPath string, exporter Exporter) (ExportedPayload, error) {
	if exporter == nil {
		return ExportedPayload{}, fmt.Errorf("exporter is required")
	}
	data, err := d.ExtractChartDataByPath(chartPath)
	if err != nil {
		return ExportedPayload{}, err
	}
	payload, err := exporter.Export(data)
	if err != nil {
		return ExportedPayload{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_CELL_PARSE_ERROR",
			message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
			err:     err,
			context: map[string]string{"chart": chartPath, "error": err.Error()},
		})
	}
	return payload, nil
}

func (d *Document) ExportAllCharts(exporter Exporter) ([]ExportedPayload, error) {
	if exporter == nil {
		return nil, fmt.Errorf("exporter is required")
	}
	charts, err := d.ExtractAllCharts()
	if err != nil {
		return nil, err
	}
	payloads := make([]ExportedPayload, 0, len(charts))
	for _, chart := range charts {
		payload, err := exporter.Export(chart)
		if err != nil {
			if d.opts.Mode == BestEffort {
				_ = d.handleExtractError(extractIssue{
					code:    "EXTRACT_CELL_PARSE_ERROR",
					message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
					err:     err,
					context: map[string]string{"chart": chart.Meta.ChartPath, "error": err.Error()},
				})
				continue
			}
			return nil, err
		}
		payloads = append(payloads, payload)
	}
	if len(payloads) == 0 {
		return []ExportedPayload{}, nil
	}
	return payloads, nil
}

func (d *Document) ExportChartByPathFormat(chartPath string, format ExportFormat) (ExportedPayload, error) {
	exporter, err := d.exporterForFormat(format)
	if err != nil {
		return ExportedPayload{}, err
	}
	return d.ExportChartByPath(chartPath, exporter)
}

func (d *Document) ExportAllChartsFormat(format ExportFormat) ([]ExportedPayload, error) {
	exporter, err := d.exporterForFormat(format)
	if err != nil {
		return nil, err
	}
	return d.ExportAllCharts(exporter)
}

type extractIssue struct {
	code    string
	message string
	err     error
	context map[string]string
}

func (d *Document) handleExtractError(issue extractIssue) error {
	if issue.code == "" {
		return issue.err
	}
	if d.opts.Mode == BestEffort {
		d.addAlert(Alert{
			Level:   "warn",
			Code:    issue.code,
			Message: issue.message,
			Context: issue.context,
		})
	}
	return issue.err
}

func (d *Document) exporterForFormat(format ExportFormat) (Exporter, error) {
	if d == nil || d.pkg == nil {
		return nil, fmt.Errorf("document not initialized")
	}
	if format == "" {
		return nil, fmt.Errorf("export format is required")
	}
	if d.exporters == nil {
		d.exporters = defaultExporterRegistry(d.opts)
	}

	exporter, ok := d.exporters.Get(format)
	if !ok || exporter == nil {
		err := fmt.Errorf("export format %q not registered", format)
		return nil, d.handleExtractError(extractIssue{
			code:    "EXPORT_FORMAT_UNSUPPORTED",
			message: extractMessageForCode("EXPORT_FORMAT_UNSUPPORTED"),
			err:     err,
			context: map[string]string{"format": string(format)},
		})
	}
	return exporter, nil
}

func (d *Document) extractChartData(chart chartdiscover.EmbeddedChart) (ExtractedChartData, error) {
	chartXML, err := d.pkg.ReadPart(chart.ChartPath)
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "CHART_DEPENDENCIES_PARSE_FAILED",
			message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
			err:     fmt.Errorf("read chart %q: %w", chart.ChartPath, err),
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}

	info, err := chartxml.ParseInfo(bytes.NewReader(chartXML))
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "CHART_DEPENDENCIES_PARSE_FAILED",
			message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}
	if info.ChartType == "mixed" {
		return d.extractMixedChartData(chart, chartXML)
	}

	deps, err := d.extractChartDependencies(EmbeddedChart{
		SlidePath:    chart.SlidePath,
		ChartPath:    chart.ChartPath,
		WorkbookPath: chart.WorkbookPath,
	})
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "CHART_DEPENDENCIES_PARSE_FAILED",
			message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}

	if deps.ChartType != "bar" && deps.ChartType != "line" && deps.ChartType != "pie" && deps.ChartType != "area" {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "CHART_TYPE_UNSUPPORTED",
			message: extractMessageForCode("CHART_TYPE_UNSUPPORTED"),
			err:     fmt.Errorf("unsupported chart type %q", deps.ChartType),
			context: map[string]string{"chart": chart.ChartPath, "slide": chart.SlidePath, "workbook": chart.WorkbookPath, "chartType": deps.ChartType},
		})
	}

	if err := validatePlanRanges(deps.Ranges); err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_INVALID_RANGE",
			message: extractMessageForCode("EXTRACT_INVALID_RANGE"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}

	wbBytes, err := d.pkg.ReadPart(chart.WorkbookPath)
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_CELL_PARSE_ERROR",
			message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
			err:     fmt.Errorf("read workbook %q: %w", chart.WorkbookPath, err),
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}

	sharedFound, sheetPath, cellRef, err := detectSharedStrings(wbBytes)
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_CELL_PARSE_ERROR",
			message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}
	if sharedFound {
		ctx := map[string]string{
			"chart":    chart.ChartPath,
			"slide":    chart.SlidePath,
			"workbook": chart.WorkbookPath,
		}
		if sheetPath != "" {
			ctx["sheetPath"] = sheetPath
		}
		if cellRef != "" {
			ctx["cell"] = cellRef
		}
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_SHAREDSTRINGS_UNSUPPORTED",
			message: extractMessageForCode("EXTRACT_SHAREDSTRINGS_UNSUPPORTED"),
			err:     fmt.Errorf("sharedStrings not supported"),
			context: ctx,
		})
	}

	wb, err := xlsxembed.Open(wbBytes)
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_CELL_PARSE_ERROR",
			message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}

	catRange, valuesRanges, nameRanges := splitDependencies(deps.Ranges)
	if deps.ChartType == "pie" {
		if len(valuesRanges) == 0 || len(valuesRanges) > 1 {
			return ExtractedChartData{}, d.handleExtractError(extractIssue{
				code:    "EXTRACT_INVALID_RANGE",
				message: extractMessageForCode("EXTRACT_INVALID_RANGE"),
				err:     fmt.Errorf("pie chart requires exactly one series"),
				context: map[string]string{
					"chart":    chart.ChartPath,
					"slide":    chart.SlidePath,
					"workbook": chart.WorkbookPath,
				},
			})
		}
	}
	labels := []string{}
	primarySheet := ""
	if catRange != nil {
		labels, err = wb.GetRangeValues(catRange.Sheet, catRange.StartCell, catRange.EndCell, xlsxembed.MissingNumericEmpty)
		if err != nil {
			return ExtractedChartData{}, d.handleWorkbookRangeError(chart, catRange.Sheet, err)
		}
		primarySheet = catRange.Sheet
	}
	if primarySheet == "" && len(valuesRanges) > 0 {
		if valueRange, ok := valuesRanges[sortedKeys(valuesRanges)[0]]; ok {
			primarySheet = valueRange.Sheet
		}
	}

	series := make([]ExtractedSeries, 0, len(valuesRanges))
	for _, index := range sortedKeys(valuesRanges) {
		valueRange := valuesRanges[index]
		values, err := wb.GetRangeValues(valueRange.Sheet, valueRange.StartCell, valueRange.EndCell, xlsxembed.MissingNumericEmpty)
		if err != nil {
			return ExtractedChartData{}, d.handleWorkbookRangeError(chart, valueRange.Sheet, err)
		}

		name := fmt.Sprintf("Series %d", index+1)
		if nameRange, ok := nameRanges[index]; ok {
			names, err := wb.GetRangeValues(nameRange.Sheet, nameRange.StartCell, nameRange.EndCell, xlsxembed.MissingNumericEmpty)
			if err != nil {
				return ExtractedChartData{}, d.handleWorkbookRangeError(chart, nameRange.Sheet, err)
			}
			if len(names) > 0 {
				trimmed := strings.TrimSpace(names[0])
				if trimmed != "" {
					name = trimmed
				}
			}
		}

		series = append(series, ExtractedSeries{
			Index: index,
			Name:  name,
			Data:  values,
		})
	}

	meta := ExtractMeta{
		ChartPath:    chart.ChartPath,
		SlidePath:    chart.SlidePath,
		WorkbookPath: chart.WorkbookPath,
		Sheet:        primarySheet,
	}

	return ExtractedChartData{
		Type:   deps.ChartType,
		Labels: labels,
		Series: series,
		Meta:   meta,
	}, nil
}

func (d *Document) extractMixedChartData(chart chartdiscover.EmbeddedChart, chartXML []byte) (ExtractedChartData, error) {
	parsed, err := chartxml.ParseMixed(bytes.NewReader(chartXML))
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_MIXED_CHART_DETECTED",
			message: extractMessageForCode("EXTRACT_MIXED_CHART_DETECTED"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}
	if len(parsed.Series) == 0 {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_MIXED_CHART_DETECTED",
			message: extractMessageForCode("EXTRACT_MIXED_CHART_DETECTED"),
			err:     fmt.Errorf("mixed chart has no series"),
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
			},
		})
	}

	seriesRanges := make(map[int]*mixedSeriesRanges, len(parsed.Series))
	for _, series := range parsed.Series {
		seriesRanges[series.Index] = &mixedSeriesRanges{series: series}
		for _, formula := range series.Formulas {
			ref, err := xlref.ParseA1Range(formula.Formula)
			if err != nil {
				return ExtractedChartData{}, d.handleExtractError(extractIssue{
					code:    "CHART_DEPENDENCIES_PARSE_FAILED",
					message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
					err:     err,
					context: map[string]string{
						"chart":    chart.ChartPath,
						"slide":    chart.SlidePath,
						"workbook": chart.WorkbookPath,
						"error":    err.Error(),
					},
				})
			}

			r := ChartRange{
				Kind:        ChartRangeKind(formula.Kind),
				SeriesIndex: series.Index,
				Sheet:       ref.Sheet,
				StartCell:   ref.StartCell,
				EndCell:     ref.EndCell,
				Formula:     formula.Formula,
			}
			entry := seriesRanges[series.Index]
			switch r.Kind {
			case RangeCategories:
				if entry.categories != nil {
					return ExtractedChartData{}, d.handleExtractError(extractIssue{
						code:    "CHART_DEPENDENCIES_PARSE_FAILED",
						message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
						err:     fmt.Errorf("duplicate categories range for series %d", series.Index),
						context: map[string]string{
							"chart":    chart.ChartPath,
							"slide":    chart.SlidePath,
							"workbook": chart.WorkbookPath,
						},
					})
				}
				entry.categories = &r
			case RangeValues:
				if entry.values != nil {
					return ExtractedChartData{}, d.handleExtractError(extractIssue{
						code:    "CHART_DEPENDENCIES_PARSE_FAILED",
						message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
						err:     fmt.Errorf("duplicate values range for series %d", series.Index),
						context: map[string]string{
							"chart":    chart.ChartPath,
							"slide":    chart.SlidePath,
							"workbook": chart.WorkbookPath,
						},
					})
				}
				entry.values = &r
			case RangeSeriesName:
				if entry.name != nil {
					return ExtractedChartData{}, d.handleExtractError(extractIssue{
						code:    "CHART_DEPENDENCIES_PARSE_FAILED",
						message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
						err:     fmt.Errorf("duplicate series name range for series %d", series.Index),
						context: map[string]string{
							"chart":    chart.ChartPath,
							"slide":    chart.SlidePath,
							"workbook": chart.WorkbookPath,
						},
					})
				}
				entry.name = &r
			}
		}
	}

	var catKey string
	for _, entry := range seriesRanges {
		if entry.categories == nil || entry.values == nil {
			return ExtractedChartData{}, d.handleExtractError(extractIssue{
				code:    "CHART_DEPENDENCIES_PARSE_FAILED",
				message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
				err:     fmt.Errorf("mixed chart requires categories and values for each series"),
				context: map[string]string{
					"chart":    chart.ChartPath,
					"slide":    chart.SlidePath,
					"workbook": chart.WorkbookPath,
				},
			})
		}
		key := entry.categories.Sheet + "!" + entry.categories.StartCell + ":" + entry.categories.EndCell
		if catKey == "" {
			catKey = key
		} else if key != catKey {
			return ExtractedChartData{}, d.handleExtractError(extractIssue{
				code:    "CHART_DEPENDENCIES_PARSE_FAILED",
				message: extractMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
				err:     fmt.Errorf("mixed chart categories must match across series"),
				context: map[string]string{
					"chart":    chart.ChartPath,
					"slide":    chart.SlidePath,
					"workbook": chart.WorkbookPath,
				},
			})
		}

		if _, err := expandRangeCells(entry.categories.StartCell, entry.categories.EndCell); err != nil {
			return ExtractedChartData{}, d.handleExtractError(extractIssue{
				code:    "EXTRACT_INVALID_RANGE",
				message: extractMessageForCode("EXTRACT_INVALID_RANGE"),
				err:     err,
				context: map[string]string{
					"chart":    chart.ChartPath,
					"slide":    chart.SlidePath,
					"workbook": chart.WorkbookPath,
					"error":    err.Error(),
				},
			})
		}
		if _, err := expandRangeCells(entry.values.StartCell, entry.values.EndCell); err != nil {
			return ExtractedChartData{}, d.handleExtractError(extractIssue{
				code:    "EXTRACT_INVALID_RANGE",
				message: extractMessageForCode("EXTRACT_INVALID_RANGE"),
				err:     err,
				context: map[string]string{
					"chart":    chart.ChartPath,
					"slide":    chart.SlidePath,
					"workbook": chart.WorkbookPath,
					"error":    err.Error(),
				},
			})
		}
	}

	wbBytes, err := d.pkg.ReadPart(chart.WorkbookPath)
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_CELL_PARSE_ERROR",
			message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
			err:     fmt.Errorf("read workbook %q: %w", chart.WorkbookPath, err),
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}

	sharedFound, sheetPath, cellRef, err := detectSharedStrings(wbBytes)
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_CELL_PARSE_ERROR",
			message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}
	if sharedFound {
		ctx := map[string]string{
			"chart":    chart.ChartPath,
			"slide":    chart.SlidePath,
			"workbook": chart.WorkbookPath,
		}
		if sheetPath != "" {
			ctx["sheetPath"] = sheetPath
		}
		if cellRef != "" {
			ctx["cell"] = cellRef
		}
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_SHAREDSTRINGS_UNSUPPORTED",
			message: extractMessageForCode("EXTRACT_SHAREDSTRINGS_UNSUPPORTED"),
			err:     fmt.Errorf("sharedStrings not supported"),
			context: ctx,
		})
	}

	wb, err := xlsxembed.Open(wbBytes)
	if err != nil {
		return ExtractedChartData{}, d.handleExtractError(extractIssue{
			code:    "EXTRACT_CELL_PARSE_ERROR",
			message: extractMessageForCode("EXTRACT_CELL_PARSE_ERROR"),
			err:     err,
			context: map[string]string{
				"chart":    chart.ChartPath,
				"slide":    chart.SlidePath,
				"workbook": chart.WorkbookPath,
				"error":    err.Error(),
			},
		})
	}

	seriesKeys := make([]int, 0, len(seriesRanges))
	for idx := range seriesRanges {
		seriesKeys = append(seriesKeys, idx)
	}
	sort.Ints(seriesKeys)

	catRange := seriesRanges[seriesKeys[0]].categories
	labels, err := wb.GetRangeValues(catRange.Sheet, catRange.StartCell, catRange.EndCell, xlsxembed.MissingNumericEmpty)
	if err != nil {
		return ExtractedChartData{}, d.handleWorkbookRangeError(chart, catRange.Sheet, err)
	}

	series := make([]ExtractedSeries, 0, len(seriesKeys))
	for _, idx := range seriesKeys {
		entry := seriesRanges[idx]
		values, err := wb.GetRangeValues(entry.values.Sheet, entry.values.StartCell, entry.values.EndCell, xlsxembed.MissingNumericEmpty)
		if err != nil {
			return ExtractedChartData{}, d.handleWorkbookRangeError(chart, entry.values.Sheet, err)
		}

		name := fmt.Sprintf("Series %d", idx+1)
		if entry.name != nil {
			names, err := wb.GetRangeValues(entry.name.Sheet, entry.name.StartCell, entry.name.EndCell, xlsxembed.MissingNumericEmpty)
			if err != nil {
				return ExtractedChartData{}, d.handleWorkbookRangeError(chart, entry.name.Sheet, err)
			}
			if len(names) > 0 {
				trimmed := strings.TrimSpace(names[0])
				if trimmed != "" {
					name = trimmed
				}
			}
		}

		series = append(series, ExtractedSeries{
			Index:    idx,
			Name:     name,
			Data:     values,
			PlotType: entry.series.PlotType,
			Axis:     entry.series.Axis,
		})
	}

	meta := ExtractMeta{
		ChartPath:    chart.ChartPath,
		SlidePath:    chart.SlidePath,
		WorkbookPath: chart.WorkbookPath,
		Sheet:        catRange.Sheet,
	}

	return ExtractedChartData{
		Type:   "mixed",
		Labels: labels,
		Series: series,
		Meta:   meta,
	}, nil
}

func (d *Document) handleWorkbookRangeError(chart chartdiscover.EmbeddedChart, sheet string, err error) error {
	code := "EXTRACT_CELL_PARSE_ERROR"
	if strings.Contains(err.Error(), "sheet") && strings.Contains(err.Error(), "not found") {
		code = "EXTRACT_SHEET_NOT_FOUND"
	}
	return d.handleExtractError(extractIssue{
		code:    code,
		message: extractMessageForCode(code),
		err:     err,
		context: map[string]string{
			"chart":    chart.ChartPath,
			"slide":    chart.SlidePath,
			"workbook": chart.WorkbookPath,
			"sheet":    sheet,
			"error":    err.Error(),
		},
	})
}

func splitDependencies(ranges []Range) (*Range, map[int]Range, map[int]Range) {
	var catRange *Range
	values := make(map[int]Range)
	names := make(map[int]Range)

	for i := range ranges {
		r := ranges[i]
		switch r.Kind {
		case RangeCategories:
			if catRange == nil || r.SeriesIndex < catRange.SeriesIndex {
				copy := r
				catRange = &copy
			}
		case RangeValues:
			if _, ok := values[r.SeriesIndex]; !ok {
				values[r.SeriesIndex] = r
			}
		case RangeSeriesName:
			if _, ok := names[r.SeriesIndex]; !ok {
				names[r.SeriesIndex] = r
			}
		}
	}

	return catRange, values, names
}

func sortedKeys(ranges map[int]Range) []int {
	keys := make([]int, 0, len(ranges))
	for key := range ranges {
		keys = append(keys, key)
	}
	sort.Ints(keys)
	return keys
}

func detectSharedStrings(data []byte) (bool, string, string, error) {
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return false, "", "", err
	}

	for _, part := range reader.File {
		if part.Name == "xl/sharedStrings.xml" {
			return true, part.Name, "", nil
		}
	}

	for _, part := range reader.File {
		if !strings.HasPrefix(part.Name, "xl/worksheets/") || !strings.HasSuffix(part.Name, ".xml") {
			continue
		}
		rc, err := part.Open()
		if err != nil {
			return false, part.Name, "", err
		}
		found, cellRef, err := scanSharedStringCells(rc)
		_ = rc.Close()
		if err != nil {
			return false, part.Name, "", err
		}
		if found {
			return true, part.Name, cellRef, nil
		}
	}

	return false, "", "", nil
}

func scanSharedStringCells(r io.Reader) (bool, string, error) {
	decoder := xml.NewDecoder(r)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return false, "", nil
		}
		if err != nil {
			return false, "", err
		}

		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "c" {
			continue
		}

		cellType := ""
		cellRef := ""
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "t":
				cellType = attr.Value
			case "r":
				cellRef = attr.Value
			}
		}
		if cellType == "s" {
			return true, cellRef, nil
		}
	}
}

func mapSkipReasonCode(skip chartdiscover.SkippedChart) string {
	switch skip.Reason {
	case chartdiscover.ReasonLinked:
		return "CHART_LINKED_WORKBOOK"
	case chartdiscover.ReasonRelsMissing:
		return "CHART_RELS_MISSING"
	case chartdiscover.ReasonWorkbookNotFound:
		return "CHART_WORKBOOK_NOT_FOUND"
	case chartdiscover.ReasonUnsupported:
		return "CHART_WORKBOOK_UNSUPPORTED_TARGET"
	default:
		return ""
	}
}

func extractSkipContext(skip chartdiscover.SkippedChart) map[string]string {
	ctx := map[string]string{
		"slide": skip.SlidePath,
		"chart": skip.ChartPath,
	}
	switch skip.Reason {
	case chartdiscover.ReasonLinked:
		ctx["target"] = skip.Target
	case chartdiscover.ReasonRelsMissing:
		ctx["rels_path"] = skip.RelsPath
	case chartdiscover.ReasonUnsupported:
		ctx["target"] = skip.Target
	}
	return ctx
}

func extractMessageForCode(code string) string {
	switch code {
	case "CHART_LINKED_WORKBOOK":
		return "Chart uses linked workbook and is skipped"
	case "CHART_RELS_MISSING":
		return "Chart relationships file is missing; chart is skipped"
	case "CHART_WORKBOOK_NOT_FOUND":
		return "No workbook relationship found for chart; chart is skipped"
	case "CHART_WORKBOOK_UNSUPPORTED_TARGET":
		return "Chart workbook target is unsupported; chart is skipped"
	case "CHART_DEPENDENCIES_PARSE_FAILED":
		return "Failed to extract chart dependencies; chart is skipped"
	case "CHART_TYPE_UNSUPPORTED":
		return "Chart type is unsupported; chart is skipped"
	case "EXTRACT_INVALID_RANGE":
		return "Chart range is invalid or unsupported; chart is skipped"
	case "EXTRACT_SHAREDSTRINGS_UNSUPPORTED":
		return "Workbook uses sharedStrings, which is unsupported"
	case "EXTRACT_SHEET_NOT_FOUND":
		return "Workbook sheet not found"
	case "EXTRACT_CELL_PARSE_ERROR":
		return "Failed to parse workbook cells"
	case "EXTRACT_MIXED_CHART_DETECTED":
		return "Mixed chart type is unsupported; chart is skipped"
	case "EXPORT_FORMAT_UNSUPPORTED":
		return "Export format is not registered"
	default:
		return "Chart extraction failed"
	}
}
