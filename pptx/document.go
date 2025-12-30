package pptx

import (
	"bytes"
	"fmt"

	"why-pptx/internal/chartcache"
	"why-pptx/internal/chartdiscover"
	"why-pptx/internal/chartxml"
	"why-pptx/internal/ooxmlpkg"
	"why-pptx/internal/xlref"
	"why-pptx/internal/xlsxembed"
)

type Alert struct {
	Level   string
	Code    string
	Message string
	Context map[string]string
}

type Logger interface {
	Debug(msg string, kv ...any)
	Info(msg string, kv ...any)
	Warn(msg string, kv ...any)
	Error(msg string, kv ...any)
}

type Option func(*Document)

type Document struct {
	pkg     *ooxmlpkg.Package
	alerts  []Alert
	logger  Logger
	strict  bool
	errMode ErrorMode
}

type EmbeddedChart struct {
	SlidePath    string
	ChartPath    string
	WorkbookPath string
}

type ChartRangeKind string

const (
	RangeCategories ChartRangeKind = "categories"
	RangeValues     ChartRangeKind = "values"
	RangeSeriesName ChartRangeKind = "seriesName"
)

type ChartRange struct {
	Kind        ChartRangeKind
	SeriesIndex int
	Sheet       string
	StartCell   string
	EndCell     string
	Formula     string
}

type ChartDependencies struct {
	SlidePath    string
	ChartPath    string
	WorkbookPath string
	ChartType    string
	Ranges       []ChartRange
}

type CellValue struct {
	Number *float64
	String *string
}

func Num(value float64) CellValue {
	return CellValue{Number: &value}
}

func Str(value string) CellValue {
	return CellValue{String: &value}
}

type CellUpdate struct {
	WorkbookPath string
	Sheet        string
	Cell         string
	Value        CellValue
}

type ErrorMode int

const (
	Strict ErrorMode = iota
	BestEffort
)

func OpenFile(path string, opts ...Option) (*Document, error) {
	pkg, err := ooxmlpkg.OpenFile(path)
	if err != nil {
		return nil, err
	}

	doc := &Document{
		pkg:     pkg,
		logger:  noopLogger{},
		errMode: Strict,
	}
	for _, opt := range opts {
		if opt != nil {
			opt(doc)
		}
	}

	return doc, nil
}

func (d *Document) SaveFile(path string) error {
	return d.pkg.SaveFile(path)
}

func (d *Document) GetChartDependencies() ([]ChartDependencies, error) {
	if d == nil || d.pkg == nil {
		return nil, fmt.Errorf("document not initialized")
	}

	charts, err := d.DiscoverEmbeddedCharts()
	if err != nil {
		return nil, err
	}

	deps := make([]ChartDependencies, 0, len(charts))
	for _, chart := range charts {
		dep, err := d.extractChartDependencies(chart)
		if err != nil {
			if d.errMode == BestEffort {
				d.addAlert(Alert{
					Level:   "warn",
					Code:    "CHART_DEPENDENCIES_PARSE_FAILED",
					Message: "Failed to extract chart dependencies; chart is skipped",
					Context: map[string]string{
						"slide":    chart.SlidePath,
						"chart":    chart.ChartPath,
						"workbook": chart.WorkbookPath,
						"error":    err.Error(),
					},
				})
				continue
			}
			return nil, err
		}
		deps = append(deps, dep)
	}

	if len(deps) == 0 {
		return []ChartDependencies{}, nil
	}
	return deps, nil
}

func (d *Document) extractChartDependencies(chart EmbeddedChart) (ChartDependencies, error) {
	data, err := d.pkg.ReadPart(chart.ChartPath)
	if err != nil {
		return ChartDependencies{}, fmt.Errorf("read chart %q: %w", chart.ChartPath, err)
	}

	parsed, err := chartxml.Parse(bytes.NewReader(data))
	if err != nil {
		return ChartDependencies{}, fmt.Errorf("parse chart %q: %w", chart.ChartPath, err)
	}

	ranges := make([]ChartRange, 0, len(parsed.Formulas))
	for _, formula := range parsed.Formulas {
		if formula.Kind != chartxml.KindCategories && formula.Kind != chartxml.KindValues && formula.Kind != chartxml.KindSeriesName {
			return ChartDependencies{}, fmt.Errorf("unknown chart formula kind %q in %s", formula.Kind, chart.ChartPath)
		}
		ref, err := xlref.ParseA1Range(formula.Formula)
		if err != nil {
			return ChartDependencies{}, fmt.Errorf("parse chart formula %q in %s: %w", formula.Formula, chart.ChartPath, err)
		}

		ranges = append(ranges, ChartRange{
			Kind:        ChartRangeKind(formula.Kind),
			SeriesIndex: formula.SeriesIndex,
			Sheet:       ref.Sheet,
			StartCell:   ref.StartCell,
			EndCell:     ref.EndCell,
			Formula:     formula.Formula,
		})
	}

	return ChartDependencies{
		SlidePath:    chart.SlidePath,
		ChartPath:    chart.ChartPath,
		WorkbookPath: chart.WorkbookPath,
		ChartType:    parsed.ChartType,
		Ranges:       ranges,
	}, nil
}

func (d *Document) DiscoverEmbeddedCharts() ([]EmbeddedChart, error) {
	if d == nil || d.pkg == nil {
		return nil, fmt.Errorf("document not initialized")
	}

	embedded, skipped, err := chartdiscover.DiscoverEmbeddedCharts(d.pkg)
	if err != nil {
		return nil, err
	}

	for _, skip := range skipped {
		switch skip.Reason {
		case chartdiscover.ReasonLinked:
			d.addAlert(Alert{
				Level:   "warn",
				Code:    "CHART_LINKED_WORKBOOK",
				Message: "Chart uses linked workbook and is skipped",
				Context: map[string]string{
					"slide":  skip.SlidePath,
					"chart":  skip.ChartPath,
					"target": skip.Target,
				},
			})
		case chartdiscover.ReasonRelsMissing:
			d.addAlert(Alert{
				Level:   "warn",
				Code:    "CHART_RELS_MISSING",
				Message: "Chart relationships file is missing; chart is skipped",
				Context: map[string]string{
					"slide":     skip.SlidePath,
					"chart":     skip.ChartPath,
					"rels_path": skip.RelsPath,
				},
			})
		case chartdiscover.ReasonWorkbookNotFound:
			d.addAlert(Alert{
				Level:   "warn",
				Code:    "CHART_WORKBOOK_NOT_FOUND",
				Message: "No workbook relationship found for chart; chart is skipped",
				Context: map[string]string{
					"slide": skip.SlidePath,
					"chart": skip.ChartPath,
				},
			})
		case chartdiscover.ReasonUnsupported:
			d.addAlert(Alert{
				Level:   "warn",
				Code:    "CHART_WORKBOOK_UNSUPPORTED_TARGET",
				Message: "Chart workbook target is unsupported; chart is skipped",
				Context: map[string]string{
					"slide":  skip.SlidePath,
					"chart":  skip.ChartPath,
					"target": skip.Target,
				},
			})
		}
	}

	if len(embedded) == 0 {
		return []EmbeddedChart{}, nil
	}

	out := make([]EmbeddedChart, len(embedded))
	for i, item := range embedded {
		out[i] = EmbeddedChart{
			SlidePath:    item.SlidePath,
			ChartPath:    item.ChartPath,
			WorkbookPath: item.WorkbookPath,
		}
	}

	return out, nil
}

func (d *Document) SetWorkbookCells(updates []CellUpdate) error {
	if d == nil || d.pkg == nil {
		return fmt.Errorf("document not initialized")
	}
	if len(updates) == 0 {
		return nil
	}

	updatesByWorkbook := make(map[string][]CellUpdate)
	for _, update := range updates {
		updatesByWorkbook[update.WorkbookPath] = append(updatesByWorkbook[update.WorkbookPath], update)
	}

	for workbookPath, wbUpdates := range updatesByWorkbook {
		if workbookPath == "" {
			if err := d.handleWorkbookUpdateError(CellUpdate{}, fmt.Errorf("workbook path is required")); err != nil {
				return err
			}
			continue
		}

		data, err := d.pkg.ReadPart(workbookPath)
		if err != nil {
			if err := d.handleWorkbookUpdateError(wbUpdates[0], fmt.Errorf("read workbook %q: %w", workbookPath, err)); err != nil {
				return err
			}
			continue
		}

		wb, err := xlsxembed.Open(data)
		if err != nil {
			if err := d.handleWorkbookUpdateError(wbUpdates[0], fmt.Errorf("open workbook %q: %w", workbookPath, err)); err != nil {
				return err
			}
			continue
		}

		applyFailed := false
		var applyErr error
		var failedUpdate CellUpdate

		for _, update := range wbUpdates {
			normalized, err := xlref.NormalizeCellRef(update.Cell)
			if err != nil {
				applyFailed = true
				applyErr = fmt.Errorf("invalid cell %q: %w", update.Cell, err)
				failedUpdate = update
				break
			}
			update.Cell = normalized

			if update.Sheet == "" {
				applyFailed = true
				applyErr = fmt.Errorf("sheet name is required")
				failedUpdate = update
				break
			}

			if err := validateCellValue(update.Value); err != nil {
				applyFailed = true
				applyErr = err
				failedUpdate = update
				break
			}

			if err := wb.SetCell(update.Sheet, update.Cell, xlsxembed.CellValue{
				Number: update.Value.Number,
				String: update.Value.String,
			}); err != nil {
				applyFailed = true
				applyErr = err
				failedUpdate = update
				break
			}
		}

		if applyFailed {
			if err := d.handleWorkbookUpdateError(failedUpdate, fmt.Errorf("update workbook %q: %w", workbookPath, applyErr)); err != nil {
				return err
			}
			continue
		}

		newBytes, err := wb.Save()
		if err != nil {
			if err := d.handleWorkbookUpdateError(wbUpdates[0], fmt.Errorf("save workbook %q: %w", workbookPath, err)); err != nil {
				return err
			}
			continue
		}

		d.pkg.WritePart(workbookPath, newBytes)
	}

	return nil
}

func (d *Document) SyncChartCaches() error {
	if d == nil || d.pkg == nil {
		return fmt.Errorf("document not initialized")
	}

	deps, err := d.GetChartDependencies()
	if err != nil {
		return err
	}
	if len(deps) == 0 {
		return nil
	}

	for _, dep := range deps {
		chartData, err := d.pkg.ReadPart(dep.ChartPath)
		if err != nil {
			if err := d.handleChartCacheError(dep, fmt.Errorf("read chart %q: %w", dep.ChartPath, err)); err != nil {
				return err
			}
			continue
		}

		wbData, err := d.pkg.ReadPart(dep.WorkbookPath)
		if err != nil {
			if err := d.handleChartCacheError(dep, fmt.Errorf("read workbook %q: %w", dep.WorkbookPath, err)); err != nil {
				return err
			}
			continue
		}

		wb, err := xlsxembed.Open(wbData)
		if err != nil {
			if err := d.handleChartCacheError(dep, fmt.Errorf("open workbook %q: %w", dep.WorkbookPath, err)); err != nil {
				return err
			}
			continue
		}

		cacheDeps, err := toCacheDeps(dep)
		if err != nil {
			if err := d.handleChartCacheError(dep, err); err != nil {
				return err
			}
			continue
		}

		updated, err := chartcache.SyncCaches(chartData, cacheDeps, func(sheet, start, end string) ([]string, error) {
			return wb.GetRangeValues(sheet, start, end)
		})
		if err != nil {
			if err := d.handleChartCacheError(dep, err); err != nil {
				return err
			}
			continue
		}

		d.pkg.WritePart(dep.ChartPath, updated)
	}

	return nil
}

// Close is a no-op in v0; Document does not hold OS resources yet.
func (d *Document) Close() error {
	return nil
}

func (d *Document) Alerts() []Alert {
	if d == nil || len(d.alerts) == 0 {
		return []Alert{}
	}

	out := make([]Alert, len(d.alerts))
	for i, alert := range d.alerts {
		out[i] = alert
		if alert.Context == nil {
			continue
		}
		ctxCopy := make(map[string]string, len(alert.Context))
		for key, value := range alert.Context {
			ctxCopy[key] = value
		}
		out[i].Context = ctxCopy
	}

	return out
}

func (d *Document) addAlert(alert Alert) {
	if d == nil {
		return
	}
	d.alerts = append(d.alerts, alert)
}

func WithLogger(logger Logger) Option {
	return func(d *Document) {
		if d == nil || logger == nil {
			return
		}
		d.logger = logger
	}
}

func WithStrict(strict bool) Option {
	return func(d *Document) {
		if d == nil {
			return
		}
		// Placeholder for future strict validation behavior.
		d.strict = strict
	}
}

func WithErrorMode(mode ErrorMode) Option {
	return func(d *Document) {
		if d == nil {
			return
		}
		d.errMode = mode
	}
}

func validateCellValue(value CellValue) error {
	if value.Number == nil && value.String == nil {
		return fmt.Errorf("cell value must specify number or string")
	}
	if value.Number != nil && value.String != nil {
		return fmt.Errorf("cell value must specify exactly one of number or string")
	}
	return nil
}

func (d *Document) handleWorkbookUpdateError(update CellUpdate, err error) error {
	if d.errMode != BestEffort {
		return err
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    "WORKBOOK_UPDATE_FAILED",
		Message: "Failed to update workbook cell; workbook is skipped",
		Context: map[string]string{
			"workbook": update.WorkbookPath,
			"sheet":    update.Sheet,
			"cell":     update.Cell,
			"error":    err.Error(),
		},
	})

	return nil
}

func toCacheDeps(dep ChartDependencies) (chartcache.Dependencies, error) {
	ranges := make([]chartcache.Range, len(dep.Ranges))
	for i, r := range dep.Ranges {
		var kind chartcache.RangeKind
		switch r.Kind {
		case RangeCategories:
			kind = chartcache.KindCategories
		case RangeValues:
			kind = chartcache.KindValues
		case RangeSeriesName:
			kind = chartcache.KindSeriesName
		default:
			return chartcache.Dependencies{}, fmt.Errorf("unsupported range kind %q", r.Kind)
		}

		ranges[i] = chartcache.Range{
			Kind:        kind,
			SeriesIndex: r.SeriesIndex,
			Sheet:       r.Sheet,
			StartCell:   r.StartCell,
			EndCell:     r.EndCell,
		}
	}

	return chartcache.Dependencies{
		ChartType: dep.ChartType,
		Ranges:    ranges,
	}, nil
}

func (d *Document) handleChartCacheError(dep ChartDependencies, err error) error {
	if d.errMode != BestEffort {
		return err
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    "CHART_CACHE_SYNC_FAILED",
		Message: "Failed to sync chart caches; chart is skipped",
		Context: map[string]string{
			"slide":    dep.SlidePath,
			"chart":    dep.ChartPath,
			"workbook": dep.WorkbookPath,
			"error":    err.Error(),
		},
	})

	return nil
}

type noopLogger struct{}

func (noopLogger) Debug(msg string, kv ...any) {}
func (noopLogger) Info(msg string, kv ...any)  {}
func (noopLogger) Warn(msg string, kv ...any)  {}
func (noopLogger) Error(msg string, kv ...any) {}
