package pptx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"

	"why-pptx/internal/chartcache"
	"why-pptx/internal/chartdiscover"
	"why-pptx/internal/chartxml"
	"why-pptx/internal/ooxmlpkg"
	"why-pptx/internal/overlaystage"
	"why-pptx/internal/postflight"
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
	pkg       *ooxmlpkg.Package
	overlay   overlaystage.Overlay
	alerts    []Alert
	logger    Logger
	strict    bool
	opts      Options
	exporters *ExporterRegistry
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

type mixedWriteSeries struct {
	SeriesIndex int
	PlotIndex   int
	PlotType    string
	Categories  ChartRange
	Values      ChartRange
	Name        *ChartRange
}

type mixedAxisGroup struct {
	CatAxID string
	ValAxID string
}

type mixedPlotBinding struct {
	PlotType  string
	AxisGroup mixedAxisGroup
	AxisRole  string
}

type mixedWriteDeps struct {
	Categories ChartRange
	Series     []mixedWriteSeries
	Bindings   []mixedPlotBinding
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

type Options struct {
	Mode     ErrorMode
	Chart    ChartOptions
	Workbook WorkbookOptions
}

type ChartOptions struct {
	CacheSync bool
}

type WorkbookOptions struct {
	MissingNumericPolicy MissingNumericPolicy
}

type MissingNumericPolicy int

const (
	MissingNumericEmpty MissingNumericPolicy = iota
	MissingNumericZero
)

type ErrorMode int

const (
	Strict ErrorMode = iota
	BestEffort
)

// DefaultOptions returns stable defaults for production use:
// Mode=Strict, Chart.CacheSync=true, Workbook.MissingNumericPolicy=MissingNumericEmpty.
func DefaultOptions() Options {
	return Options{
		Mode:  Strict,
		Chart: ChartOptions{CacheSync: true},
		Workbook: WorkbookOptions{
			MissingNumericPolicy: MissingNumericEmpty,
		},
	}
}

func OpenFile(path string, opts ...Option) (*Document, error) {
	pkg, err := ooxmlpkg.OpenFile(path)
	if err != nil {
		return nil, err
	}

	overlay, err := overlaystage.NewPackageOverlay(pkg)
	if err != nil {
		return nil, err
	}

	doc := &Document{
		pkg:     pkg,
		overlay: overlay,
		logger:  noopLogger{},
		opts:    DefaultOptions(),
	}
	for _, opt := range opts {
		if opt != nil {
			opt(doc)
		}
	}
	if doc.exporters == nil {
		doc.exporters = defaultExporterRegistry(doc.opts)
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
			if d.opts.Mode == BestEffort {
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

func (d *Document) setWorkbookCellsInOverlay(overlay overlaystage.Overlay, updates []CellUpdate) error {
	if d == nil || overlay == nil {
		return fmt.Errorf("overlay not initialized")
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
			return fmt.Errorf("workbook path is required")
		}

		data, err := overlay.Get(workbookPath)
		if err != nil {
			return fmt.Errorf("read workbook %q: %w", workbookPath, err)
		}

		wb, err := xlsxembed.Open(data)
		if err != nil {
			return fmt.Errorf("open workbook %q: %w", workbookPath, err)
		}

		for _, update := range wbUpdates {
			normalized, err := xlref.NormalizeCellRef(update.Cell)
			if err != nil {
				return fmt.Errorf("invalid cell %q: %w", update.Cell, err)
			}
			update.Cell = normalized

			if update.Sheet == "" {
				return fmt.Errorf("sheet name is required")
			}

			if err := validateCellValue(update.Value); err != nil {
				return err
			}

			if err := wb.SetCell(update.Sheet, update.Cell, xlsxembed.CellValue{
				Number: update.Value.Number,
				String: update.Value.String,
			}); err != nil {
				return err
			}
		}

		newBytes, err := wb.Save()
		if err != nil {
			return fmt.Errorf("save workbook %q: %w", workbookPath, err)
		}

		if err := overlay.Set(workbookPath, newBytes); err != nil {
			return fmt.Errorf("write workbook %q: %w", workbookPath, err)
		}
	}

	return nil
}

func (d *Document) SyncChartCaches() error {
	if d == nil || d.pkg == nil {
		return fmt.Errorf("document not initialized")
	}
	if !d.opts.Chart.CacheSync {
		return nil
	}

	deps, err := d.GetChartDependencies()
	if err != nil {
		return err
	}
	if len(deps) == 0 {
		return nil
	}

	for _, dep := range deps {
		if err := d.validateWritableChart(dep); err != nil {
			if d.opts.Mode == BestEffort {
				continue
			}
			return err
		}

		ctx := d.validateContext(dep)
		err := d.withChartStage(ctx, func(stage overlaystage.Overlay) error {
			if dep.ChartType == "mixed" {
				return d.syncMixedChartCacheInOverlay(stage, dep)
			}
			return d.syncChartCacheInOverlay(stage, dep)
		})
		if err != nil {
			if postflight.IsPostflightError(err) {
				return err
			}
			if err := d.handleChartCacheError(dep, err); err != nil {
				return err
			}
			continue
		}
	}

	return nil
}

func (d *Document) ApplyChartData(chartIndex int, data map[string][]string) error {
	if d == nil || d.pkg == nil {
		return fmt.Errorf("document not initialized")
	}
	if chartIndex < 0 {
		return fmt.Errorf("chart index out of range")
	}

	deps, err := d.GetChartDependencies()
	if err != nil {
		return err
	}
	if chartIndex >= len(deps) {
		return fmt.Errorf("chart index out of range")
	}

	dep := deps[chartIndex]
	if dep.ChartType == "mixed" {
		return d.applyMixedChartData(chartIndex, dep, data)
	}
	if err := d.validateWritableChart(dep); err != nil {
		return err
	}
	categories, hasCategories := data["categories"]
	if hasCategories {
		categoriesLen := len(categories)
		for _, r := range dep.Ranges {
			if r.Kind != RangeValues {
				continue
			}
			key := fmt.Sprintf("values:%d", r.SeriesIndex)
			values, ok := data[key]
			if !ok {
				continue
			}
			if len(values) != categoriesLen {
				return d.handleChartDataMismatch(chartIndex, categoriesLen, len(values), r.SeriesIndex)
			}
		}
	}
	updates := make([]CellUpdate, 0)

	for _, r := range dep.Ranges {
		switch r.Kind {
		case RangeCategories:
			if !hasCategories {
				return fmt.Errorf("categories data is required")
			}
			cells, err := expandRangeCells(r.StartCell, r.EndCell)
			if err != nil {
				return err
			}
			if len(categories) != len(cells) {
				return fmt.Errorf("categories length mismatch: expected %d got %d", len(cells), len(categories))
			}
			for i, cell := range cells {
				updates = append(updates, CellUpdate{
					WorkbookPath: dep.WorkbookPath,
					Sheet:        r.Sheet,
					Cell:         cell,
					Value:        Str(categories[i]),
				})
			}
		case RangeValues:
			key := fmt.Sprintf("values:%d", r.SeriesIndex)
			values, ok := data[key]
			if !ok {
				return fmt.Errorf("values data missing for series %d", r.SeriesIndex)
			}
			cells, err := expandRangeCells(r.StartCell, r.EndCell)
			if err != nil {
				return err
			}
			if len(values) != len(cells) {
				return fmt.Errorf("values length mismatch for series %d: expected %d got %d", r.SeriesIndex, len(cells), len(values))
			}
			for i, cell := range cells {
				value := strings.TrimSpace(values[i])
				number, err := strconv.ParseFloat(value, 64)
				if err != nil {
					return fmt.Errorf("invalid numeric value %q for series %d", values[i], r.SeriesIndex)
				}
				updates = append(updates, CellUpdate{
					WorkbookPath: dep.WorkbookPath,
					Sheet:        r.Sheet,
					Cell:         cell,
					Value:        Num(number),
				})
			}
		}
	}

	if len(updates) == 0 {
		return fmt.Errorf("no chart ranges matched")
	}

	ctx := d.validateContext(dep)
	return d.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		if err := d.setWorkbookCellsInOverlay(stage, updates); err != nil {
			return err
		}
		if d.opts.Chart.CacheSync {
			return d.syncChartCacheInOverlay(stage, dep)
		}
		return nil
	})
}

func (d *Document) applyMixedChartData(chartIndex int, dep ChartDependencies, data map[string][]string) error {
	mixedDeps, code, err := d.mixedWriteDependencies(dep)
	if err != nil {
		return d.handleMixedWriteError(dep, code, err)
	}

	categories, hasCategories := data["categories"]
	if !hasCategories {
		return fmt.Errorf("categories data is required")
	}

	valuesBySeries := make([][]string, len(mixedDeps.Series))
	for i := range mixedDeps.Series {
		key := fmt.Sprintf("values:%d", i)
		values, ok := data[key]
		if !ok {
			return fmt.Errorf("values data missing for series %d", i)
		}
		if len(values) != len(categories) {
			if err := d.handleChartDataMismatch(chartIndex, len(categories), len(values), i); err != nil {
				return err
			}
			return nil
		}
		valuesBySeries[i] = values
	}

	updates := make([]CellUpdate, 0)
	catCells, err := expandRangeCells(mixedDeps.Categories.StartCell, mixedDeps.Categories.EndCell)
	if err != nil {
		return err
	}
	if len(categories) != len(catCells) {
		return fmt.Errorf("categories length mismatch: expected %d got %d", len(catCells), len(categories))
	}
	for i, cell := range catCells {
		updates = append(updates, CellUpdate{
			WorkbookPath: dep.WorkbookPath,
			Sheet:        mixedDeps.Categories.Sheet,
			Cell:         cell,
			Value:        Str(categories[i]),
		})
	}

	for i, series := range mixedDeps.Series {
		values := valuesBySeries[i]
		cells, err := expandRangeCells(series.Values.StartCell, series.Values.EndCell)
		if err != nil {
			return err
		}
		if len(values) != len(cells) {
			return fmt.Errorf("values length mismatch for series %d: expected %d got %d", i, len(cells), len(values))
		}
		for j, cell := range cells {
			value := strings.TrimSpace(values[j])
			number, err := strconv.ParseFloat(value, 64)
			if err != nil {
				return fmt.Errorf("invalid numeric value %q for series %d", values[j], i)
			}
			updates = append(updates, CellUpdate{
				WorkbookPath: dep.WorkbookPath,
				Sheet:        series.Values.Sheet,
				Cell:         cell,
				Value:        Num(number),
			})
		}
	}

	if len(updates) == 0 {
		return fmt.Errorf("no chart ranges matched")
	}

	ctx := d.validateContext(dep)
	return d.withChartStage(ctx, func(stage overlaystage.Overlay) error {
		if err := d.setWorkbookCellsInOverlay(stage, updates); err != nil {
			return err
		}
		if d.opts.Chart.CacheSync {
			return d.syncMixedChartCacheInOverlay(stage, dep)
		}
		return nil
	})
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

func (d *Document) HasAlerts() bool {
	if d == nil {
		return false
	}
	return len(d.alerts) > 0
}

func (d *Document) AlertsByCode(code string) []Alert {
	if d == nil || code == "" {
		return []Alert{}
	}

	filtered := make([]Alert, 0)
	for _, alert := range d.alerts {
		if alert.Code != code {
			continue
		}
		copyAlert := alert
		if alert.Context != nil {
			ctxCopy := make(map[string]string, len(alert.Context))
			for key, value := range alert.Context {
				ctxCopy[key] = value
			}
			copyAlert.Context = ctxCopy
		}
		filtered = append(filtered, copyAlert)
	}

	if len(filtered) == 0 {
		return []Alert{}
	}
	return filtered
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

// WithExporterRegistry overrides the default exporter registry.
func WithExporterRegistry(registry *ExporterRegistry) Option {
	return func(d *Document) {
		if d == nil || registry == nil {
			return
		}
		d.exporters = registry
	}
}

// WithOptions replaces the document options. Use DefaultOptions() as a base.
func WithOptions(opts Options) Option {
	return func(d *Document) {
		if d == nil {
			return
		}
		d.opts = opts
	}
}

// Deprecated: use WithOptions and Options.Mode instead.
func WithBestEffort(bestEffort bool) Option {
	return func(d *Document) {
		if d == nil {
			return
		}
		if bestEffort {
			d.opts.Mode = BestEffort
		} else {
			d.opts.Mode = Strict
		}
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
		d.opts.Mode = mode
	}
}

func (d *Document) withChartStage(ctx postflight.ValidateContext, fn func(stage overlaystage.Overlay) error) error {
	if d == nil || d.pkg == nil {
		return fmt.Errorf("document not initialized")
	}
	if d.overlay == nil {
		overlay, err := overlaystage.NewPackageOverlay(d.pkg)
		if err != nil {
			return err
		}
		d.overlay = overlay
	}

	stage := overlaystage.NewStagingOverlay(d.overlay)
	if err := fn(stage); err != nil {
		stage.Discard()
		return err
	}

	validator := postflight.NewPostflightValidator(&postflight.Document{
		Overlay: d.overlay,
		EmitAlert: func(code, message string, ctx map[string]string) {
			d.addAlert(Alert{
				Level:   "error",
				Code:    code,
				Message: message,
				Context: ctx,
			})
		},
	})
	if err := validator.ValidateChartStage(ctx, stage); err != nil {
		stage.Discard()
		return err
	}

	if err := stage.Commit(); err != nil {
		stage.Discard()
		return err
	}
	return nil
}

func (d *Document) validateContext(dep ChartDependencies) postflight.ValidateContext {
	mode := postflight.ModeStrict
	if d.opts.Mode == BestEffort {
		mode = postflight.ModeBestEffort
	}
	return postflight.ValidateContext{
		ChartPath:            dep.ChartPath,
		SlidePath:            dep.SlidePath,
		WorkbookPath:         dep.WorkbookPath,
		Mode:                 mode,
		CacheSyncEnabled:     d.opts.Chart.CacheSync,
		MissingNumericPolicy: int(d.opts.Workbook.MissingNumericPolicy),
	}
}

func (d *Document) syncChartCacheInOverlay(overlay overlaystage.Overlay, dep ChartDependencies) error {
	if overlay == nil {
		return fmt.Errorf("overlay not initialized")
	}

	chartData, err := overlay.Get(dep.ChartPath)
	if err != nil {
		return fmt.Errorf("read chart %q: %w", dep.ChartPath, err)
	}

	wbData, err := overlay.Get(dep.WorkbookPath)
	if err != nil {
		return fmt.Errorf("read workbook %q: %w", dep.WorkbookPath, err)
	}

	wb, err := xlsxembed.Open(wbData)
	if err != nil {
		return fmt.Errorf("open workbook %q: %w", dep.WorkbookPath, err)
	}

	cacheDeps, err := toCacheDeps(dep)
	if err != nil {
		return err
	}

	updated, err := chartcache.SyncCaches(chartData, cacheDeps, func(kind chartcache.RangeKind, sheet, start, end string) ([]string, error) {
		policy := xlsxembed.MissingNumericEmpty
		if kind == chartcache.KindValues {
			policy = xlsxembed.MissingNumericPolicy(d.opts.Workbook.MissingNumericPolicy)
		}
		return wb.GetRangeValues(sheet, start, end, policy)
	})
	if err != nil {
		return err
	}

	if err := overlay.Set(dep.ChartPath, updated); err != nil {
		return fmt.Errorf("write chart %q: %w", dep.ChartPath, err)
	}
	return nil
}

func (d *Document) syncMixedChartCacheInOverlay(overlay overlaystage.Overlay, dep ChartDependencies) error {
	if overlay == nil {
		return fmt.Errorf("overlay not initialized")
	}

	chartData, err := overlay.Get(dep.ChartPath)
	if err != nil {
		return fmt.Errorf("read chart %q: %w", dep.ChartPath, err)
	}

	wbData, err := overlay.Get(dep.WorkbookPath)
	if err != nil {
		return fmt.Errorf("read workbook %q: %w", dep.WorkbookPath, err)
	}

	wb, err := xlsxembed.Open(wbData)
	if err != nil {
		return fmt.Errorf("open workbook %q: %w", dep.WorkbookPath, err)
	}

	mixedDeps, _, err := mixedWriteDependenciesFromChart(chartData)
	if err != nil {
		return err
	}

	barRanges := make([]chartcache.Range, 0)
	lineRanges := make([]chartcache.Range, 0)
	for _, series := range mixedDeps.Series {
		target := &barRanges
		if series.PlotType == "line" {
			target = &lineRanges
		}
		*target = append(*target, chartcache.Range{
			Kind:        chartcache.KindCategories,
			SeriesIndex: series.PlotIndex,
			Sheet:       series.Categories.Sheet,
			StartCell:   series.Categories.StartCell,
			EndCell:     series.Categories.EndCell,
		})
		*target = append(*target, chartcache.Range{
			Kind:        chartcache.KindValues,
			SeriesIndex: series.PlotIndex,
			Sheet:       series.Values.Sheet,
			StartCell:   series.Values.StartCell,
			EndCell:     series.Values.EndCell,
		})
		if series.Name != nil {
			*target = append(*target, chartcache.Range{
				Kind:        chartcache.KindSeriesName,
				SeriesIndex: series.PlotIndex,
				Sheet:       series.Name.Sheet,
				StartCell:   series.Name.StartCell,
				EndCell:     series.Name.EndCell,
			})
		}
	}

	if len(barRanges) == 0 || len(lineRanges) == 0 {
		return fmt.Errorf("mixed chart requires bar and line series")
	}

	provider := func(kind chartcache.RangeKind, sheet, start, end string) ([]string, error) {
		policy := xlsxembed.MissingNumericEmpty
		if kind == chartcache.KindValues {
			policy = xlsxembed.MissingNumericPolicy(d.opts.Workbook.MissingNumericPolicy)
		}
		return wb.GetRangeValues(sheet, start, end, policy)
	}

	updated, err := chartcache.SyncCaches(chartData, chartcache.Dependencies{
		ChartType: "bar",
		Ranges:    barRanges,
	}, provider)
	if err != nil {
		return err
	}

	updated, err = chartcache.SyncCaches(updated, chartcache.Dependencies{
		ChartType: "line",
		Ranges:    lineRanges,
	}, provider)
	if err != nil {
		return err
	}

	if err := overlay.Set(dep.ChartPath, updated); err != nil {
		return fmt.Errorf("write chart %q: %w", dep.ChartPath, err)
	}
	return nil
}

func (d *Document) mixedWriteDependencies(dep ChartDependencies) (*mixedWriteDeps, string, error) {
	if d == nil || d.pkg == nil {
		return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("document not initialized")
	}

	data, err := d.pkg.ReadPart(dep.ChartPath)
	if err != nil {
		return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("read chart %q: %w", dep.ChartPath, err)
	}

	return mixedWriteDependenciesFromChart(data)
}

func mixedWriteDependenciesFromChart(chartXML []byte) (*mixedWriteDeps, string, error) {
	parsed, err := chartxml.ParseMixed(bytes.NewReader(chartXML))
	if err != nil {
		code := "WRITE_MIX_UNSUPPORTED_SHAPE"
		if strings.Contains(err.Error(), "parse mixed chart") {
			code = "CHART_DEPENDENCIES_PARSE_FAILED"
		}
		return nil, code, err
	}
	if len(parsed.Series) == 0 {
		return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("mixed chart has no series")
	}

	barPlot, linePlot, err := findMixedPlots(parsed.Plots)
	if err != nil {
		return nil, "WRITE_MIX_UNSUPPORTED_SHAPE", err
	}

	if len(barPlot.AxisIDs) != 2 || len(linePlot.AxisIDs) != 2 {
		return nil, "WRITE_MIX_AXIS_GROUP_INVALID", fmt.Errorf("mixed chart requires exactly two axis ids per plot")
	}

	usesSecondaryAxis := !equalStringSlice(barPlot.AxisIDs, linePlot.AxisIDs)
	bindings := make([]mixedPlotBinding, 0, 2)
	if usesSecondaryAxis {
		secondaryBindings, err := validateSecondaryAxisGroups(barPlot, linePlot, parsed.AxisGroups)
		if err != nil {
			return nil, err.code, err.err
		}
		bindings = append(bindings, secondaryBindings...)
	} else {
		bindings = append(bindings, buildSingleAxisBindings(barPlot, linePlot, parsed.AxisGroups)...)
	}

	seriesRanges := make(map[int]*mixedWriteSeries, len(parsed.Series))
	for _, series := range parsed.Series {
		if series.PlotType != "bar" && series.PlotType != "line" {
			return nil, "WRITE_MIX_UNSUPPORTED_SHAPE", fmt.Errorf("unsupported plot type %q", series.PlotType)
		}
		entry := seriesRanges[series.Index]
		if entry == nil {
			entry = &mixedWriteSeries{
				SeriesIndex: series.Index,
				PlotIndex:   series.PlotIndex,
				PlotType:    series.PlotType,
			}
			seriesRanges[series.Index] = entry
		}
		if entry.PlotType != series.PlotType {
			return nil, "WRITE_MIX_UNSUPPORTED_SHAPE", fmt.Errorf("series %d plot type mismatch", series.Index)
		}
		if entry.PlotIndex != series.PlotIndex {
			return nil, "WRITE_MIX_UNSUPPORTED_SHAPE", fmt.Errorf("series %d plot index mismatch", series.Index)
		}

		for _, formula := range series.Formulas {
			ref, err := xlref.ParseA1Range(formula.Formula)
			if err != nil {
				return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("parse chart formula %q: %w", formula.Formula, err)
			}

			r := ChartRange{
				Kind:        ChartRangeKind(formula.Kind),
				SeriesIndex: series.Index,
				Sheet:       ref.Sheet,
				StartCell:   ref.StartCell,
				EndCell:     ref.EndCell,
				Formula:     formula.Formula,
			}

			switch r.Kind {
			case RangeCategories:
				if entry.Categories.Sheet != "" {
					return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("duplicate categories range for series %d", series.Index)
				}
				entry.Categories = r
			case RangeValues:
				if entry.Values.Sheet != "" {
					return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("duplicate values range for series %d", series.Index)
				}
				entry.Values = r
			case RangeSeriesName:
				if entry.Name != nil {
					return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("duplicate series name range for series %d", series.Index)
				}
				copy := r
				entry.Name = &copy
			default:
				return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("unknown chart formula kind %q", formula.Kind)
			}
		}
	}

	var catKey string
	var catRange ChartRange
	for _, entry := range seriesRanges {
		if entry.Categories.Sheet == "" || entry.Values.Sheet == "" {
			return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("mixed chart requires categories and values for each series")
		}
		if _, err := expandRangeCells(entry.Categories.StartCell, entry.Categories.EndCell); err != nil {
			return nil, "CHART_DEPENDENCIES_PARSE_FAILED", err
		}
		if _, err := expandRangeCells(entry.Values.StartCell, entry.Values.EndCell); err != nil {
			return nil, "CHART_DEPENDENCIES_PARSE_FAILED", err
		}
		key := entry.Categories.Sheet + "!" + entry.Categories.StartCell + ":" + entry.Categories.EndCell
		if catKey == "" {
			catKey = key
			catRange = entry.Categories
		} else if key != catKey {
			return nil, "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("mixed chart categories must match across series")
		}
	}

	ordered, err := orderMixedSeries(seriesRanges)
	if err != nil {
		return nil, "WRITE_MIX_UNSUPPORTED_SHAPE", err
	}

	return &mixedWriteDeps{
		Categories: catRange,
		Series:     ordered,
		Bindings:   bindings,
	}, "", nil
}

func findMixedPlots(plots []chartxml.MixedPlot) (chartxml.MixedPlot, chartxml.MixedPlot, error) {
	var barPlot *chartxml.MixedPlot
	var linePlot *chartxml.MixedPlot
	for i := range plots {
		plot := &plots[i]
		switch plot.PlotType {
		case "bar":
			if barPlot != nil {
				return chartxml.MixedPlot{}, chartxml.MixedPlot{}, fmt.Errorf("multiple bar plots are unsupported")
			}
			barPlot = plot
		case "line":
			if linePlot != nil {
				return chartxml.MixedPlot{}, chartxml.MixedPlot{}, fmt.Errorf("multiple line plots are unsupported")
			}
			linePlot = plot
		default:
			return chartxml.MixedPlot{}, chartxml.MixedPlot{}, fmt.Errorf("unsupported plot type %q", plot.PlotType)
		}
	}
	if barPlot == nil || linePlot == nil {
		return chartxml.MixedPlot{}, chartxml.MixedPlot{}, fmt.Errorf("mixed chart requires bar and line plots")
	}
	if len(plots) != 2 {
		return chartxml.MixedPlot{}, chartxml.MixedPlot{}, fmt.Errorf("mixed chart requires exactly two plots")
	}
	return *barPlot, *linePlot, nil
}

type axisValidationError struct {
	code string
	err  error
}

func validateSecondaryAxisGroups(barPlot, linePlot chartxml.MixedPlot, groups []chartxml.AxisGroup) ([]mixedPlotBinding, *axisValidationError) {
	if len(groups) < 2 {
		return nil, &axisValidationError{
			code: "WRITE_MIX_AXIS_GROUP_INVALID",
			err:  fmt.Errorf("secondary axis requires two axis groups"),
		}
	}
	if len(groups) > 2 {
		return nil, &axisValidationError{
			code: "WRITE_MIX_SECONDARY_AXIS_UNSUPPORTED_SHAPE",
			err:  fmt.Errorf("secondary axis supports exactly two axis groups"),
		}
	}

	barGroup, ok := axisGroupForPlot(barPlot, groups)
	if !ok {
		return nil, &axisValidationError{
			code: "WRITE_MIX_AXIS_GROUP_INVALID",
			err:  fmt.Errorf("bar plot axis group not found"),
		}
	}
	lineGroup, ok := axisGroupForPlot(linePlot, groups)
	if !ok {
		return nil, &axisValidationError{
			code: "WRITE_MIX_AXIS_GROUP_INVALID",
			err:  fmt.Errorf("line plot axis group not found"),
		}
	}

	if barGroup.CatAxID == lineGroup.CatAxID && barGroup.ValAxID == lineGroup.ValAxID {
		return nil, &axisValidationError{
			code: "WRITE_MIX_SECONDARY_AXIS_UNSUPPORTED_SHAPE",
			err:  fmt.Errorf("bar and line plots must use different axis groups"),
		}
	}

	barRole, barOK := axisRoleFromGroup(*barGroup)
	lineRole, lineOK := axisRoleFromGroup(*lineGroup)
	if !(barOK && lineOK && barRole != lineRole) {
		barRole = "primary"
		lineRole = "secondary"
	}

	return []mixedPlotBinding{
		{
			PlotType: "bar",
			AxisGroup: mixedAxisGroup{
				CatAxID: barGroup.CatAxID,
				ValAxID: barGroup.ValAxID,
			},
			AxisRole: barRole,
		},
		{
			PlotType: "line",
			AxisGroup: mixedAxisGroup{
				CatAxID: lineGroup.CatAxID,
				ValAxID: lineGroup.ValAxID,
			},
			AxisRole: lineRole,
		},
	}, nil
}

func buildSingleAxisBindings(barPlot, linePlot chartxml.MixedPlot, groups []chartxml.AxisGroup) []mixedPlotBinding {
	role := "primary"
	group, ok := axisGroupForPlot(barPlot, groups)
	if ok {
		if detected, ok := axisRoleFromGroup(*group); ok {
			role = detected
		}
	}

	axisGroup := mixedAxisGroup{}
	if group != nil {
		axisGroup = mixedAxisGroup{CatAxID: group.CatAxID, ValAxID: group.ValAxID}
	}

	return []mixedPlotBinding{
		{PlotType: "bar", AxisGroup: axisGroup, AxisRole: role},
		{PlotType: "line", AxisGroup: axisGroup, AxisRole: role},
	}
}

func axisGroupForPlot(plot chartxml.MixedPlot, groups []chartxml.AxisGroup) (*chartxml.AxisGroup, bool) {
	if len(plot.AxisIDs) != 2 {
		return nil, false
	}
	for i := range groups {
		ids := []string{groups[i].CatAxID, groups[i].ValAxID}
		sort.Strings(ids)
		if equalStringSlice(ids, plot.AxisIDs) {
			return &groups[i], true
		}
	}
	return nil, false
}

func axisRoleFromGroup(group chartxml.AxisGroup) (string, bool) {
	switch group.ValAxisPos {
	case "l", "b":
		return "primary", true
	case "r", "t":
		return "secondary", true
	}
	if group.ValHasMajorGridlines || group.ValHasMinorGridlines {
		return "primary", true
	}
	return "", false
}

func orderMixedSeries(series map[int]*mixedWriteSeries) ([]mixedWriteSeries, error) {
	bars := make([]mixedWriteSeries, 0)
	lines := make([]mixedWriteSeries, 0)

	for _, entry := range series {
		switch entry.PlotType {
		case "bar":
			bars = append(bars, *entry)
		case "line":
			lines = append(lines, *entry)
		default:
			return nil, fmt.Errorf("unsupported plot type %q", entry.PlotType)
		}
	}

	if len(bars) == 0 || len(lines) == 0 {
		return nil, fmt.Errorf("mixed chart requires bar and line series")
	}

	sort.Slice(bars, func(i, j int) bool {
		if bars[i].PlotIndex == bars[j].PlotIndex {
			return bars[i].SeriesIndex < bars[j].SeriesIndex
		}
		return bars[i].PlotIndex < bars[j].PlotIndex
	})
	sort.Slice(lines, func(i, j int) bool {
		if lines[i].PlotIndex == lines[j].PlotIndex {
			return lines[i].SeriesIndex < lines[j].SeriesIndex
		}
		return lines[i].PlotIndex < lines[j].PlotIndex
	})

	return append(bars, lines...), nil
}

func equalStringSlice(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
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
	if d.opts.Mode != BestEffort {
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
	if d.opts.Mode != BestEffort {
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

func (d *Document) validateWritableChart(dep ChartDependencies) error {
	switch dep.ChartType {
	case "mixed":
		if code, err := d.validateMixedWrite(dep); err != nil {
			return d.handleMixedWriteError(dep, code, err)
		}
	case "pie":
		if code, err := validatePieDependencies(dep); err != nil {
			return d.handlePieWriteError(dep, code, err)
		}
	case "area":
		if code, err := d.validateAreaVariant(dep); err != nil {
			return d.handleAreaWriteError(dep, code, err)
		}
		if code, err := validateAreaDependencies(dep); err != nil {
			return d.handleAreaWriteError(dep, code, err)
		}
	case "bar", "line":
		return nil
	default:
		return d.handleChartTypeUnsupported(dep)
	}
	return nil
}

func (d *Document) handleChartTypeUnsupported(dep ChartDependencies) error {
	err := fmt.Errorf("unsupported chart type %q", dep.ChartType)
	if d.opts.Mode != BestEffort {
		return err
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    "CHART_TYPE_UNSUPPORTED",
		Message: "Chart type is unsupported; chart is skipped",
		Context: map[string]string{
			"slide":     dep.SlidePath,
			"chart":     dep.ChartPath,
			"chartType": dep.ChartType,
		},
	})

	return err
}

func (d *Document) validateMixedWrite(dep ChartDependencies) (string, error) {
	_, code, err := d.mixedWriteDependencies(dep)
	return code, err
}

func (d *Document) handleMixedWriteError(dep ChartDependencies, code string, err error) error {
	if d.opts.Mode != BestEffort {
		return err
	}

	message := "Mixed chart is unsupported; chart is skipped"
	switch code {
	case "WRITE_MIX_SECONDARY_AXIS_UNSUPPORTED":
		message = "Mixed chart uses secondary axis; chart is skipped"
	case "WRITE_MIX_SECONDARY_AXIS_UNSUPPORTED_SHAPE":
		message = "Mixed chart secondary axis shape is unsupported; chart is skipped"
	case "WRITE_MIX_AXIS_GROUP_INVALID":
		message = "Mixed chart axis groups are invalid; chart is skipped"
	case "WRITE_MIX_UNSUPPORTED_SHAPE":
		message = "Mixed chart shape is unsupported; chart is skipped"
	case "CHART_DEPENDENCIES_PARSE_FAILED":
		message = "Failed to extract chart dependencies; chart is skipped"
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    code,
		Message: message,
		Context: map[string]string{
			"slide":    dep.SlidePath,
			"chart":    dep.ChartPath,
			"workbook": dep.WorkbookPath,
			"error":    err.Error(),
		},
	})

	return err
}

func (d *Document) handlePieWriteError(dep ChartDependencies, code string, err error) error {
	if d.opts.Mode != BestEffort {
		return err
	}

	message := "Pie chart is unsupported; chart is skipped"
	if code == "WRITE_PIE_MULTIPLE_SERIES_UNSUPPORTED" {
		message = "Pie charts with multiple series are unsupported; chart is skipped"
	} else if code == "CHART_DEPENDENCIES_PARSE_FAILED" {
		message = "Failed to extract chart dependencies; chart is skipped"
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    code,
		Message: message,
		Context: map[string]string{
			"slide":    dep.SlidePath,
			"chart":    dep.ChartPath,
			"workbook": dep.WorkbookPath,
			"error":    err.Error(),
		},
	})

	return err
}

func (d *Document) handleAreaWriteError(dep ChartDependencies, code string, err error) error {
	if d.opts.Mode != BestEffort {
		return err
	}

	message := "Area chart is unsupported; chart is skipped"
	if code == "WRITE_AREA_UNSUPPORTED_VARIANT" {
		message = "Area chart variant is unsupported; chart is skipped"
	} else if code == "CHART_DEPENDENCIES_PARSE_FAILED" {
		message = "Failed to extract chart dependencies; chart is skipped"
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    code,
		Message: message,
		Context: map[string]string{
			"slide":    dep.SlidePath,
			"chart":    dep.ChartPath,
			"workbook": dep.WorkbookPath,
			"error":    err.Error(),
		},
	})

	return err
}

func (d *Document) validateAreaVariant(dep ChartDependencies) (string, error) {
	if d == nil || d.pkg == nil {
		return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("document not initialized")
	}
	data, err := d.pkg.ReadPart(dep.ChartPath)
	if err != nil {
		return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("read chart %q: %w", dep.ChartPath, err)
	}

	decoder := xml.NewDecoder(bytes.NewReader(data))
	areaDepth := 0
	areaCharts := 0
	axIDs := make(map[string]struct{})

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("parse chart %q: %w", dep.ChartPath, err)
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "areaChart":
				areaDepth++
				areaCharts++
			case "area3DChart":
				return "WRITE_AREA_UNSUPPORTED_VARIANT", fmt.Errorf("area3D charts are unsupported")
			case "grouping":
				if areaDepth > 0 {
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" {
							if attr.Value == "stacked" || attr.Value == "percentStacked" {
								return "WRITE_AREA_UNSUPPORTED_VARIANT", fmt.Errorf("stacked area charts are unsupported")
							}
						}
					}
				}
			case "axId":
				if areaDepth > 0 {
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" && attr.Value != "" {
							axIDs[attr.Value] = struct{}{}
						}
					}
				}
			}
		case xml.EndElement:
			if tok.Name.Local == "areaChart" && areaDepth > 0 {
				areaDepth--
			}
		}
	}

	if areaCharts > 1 {
		return "WRITE_AREA_UNSUPPORTED_VARIANT", fmt.Errorf("multiple area charts are unsupported")
	}
	if len(axIDs) > 2 {
		return "WRITE_AREA_UNSUPPORTED_VARIANT", fmt.Errorf("secondary axis area charts are unsupported")
	}
	return "", nil
}

func validatePieDependencies(dep ChartDependencies) (string, error) {
	valueSeries := make(map[int]struct{})
	catSeries := make(map[int]struct{})

	for _, r := range dep.Ranges {
		switch r.Kind {
		case RangeValues:
			valueSeries[r.SeriesIndex] = struct{}{}
		case RangeCategories:
			catSeries[r.SeriesIndex] = struct{}{}
		}
	}

	if len(valueSeries) == 0 || len(catSeries) == 0 {
		return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("pie chart requires categories and values")
	}
	if len(valueSeries) > 1 || len(catSeries) > 1 {
		return "WRITE_PIE_MULTIPLE_SERIES_UNSUPPORTED", fmt.Errorf("pie chart requires exactly one series")
	}

	seriesIndex := -1
	for idx := range valueSeries {
		seriesIndex = idx
	}
	for idx := range catSeries {
		if idx != seriesIndex {
			return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("pie chart categories/values series mismatch")
		}
	}

	return "", nil
}

func validateAreaDependencies(dep ChartDependencies) (string, error) {
	valueRanges := make(map[int]ChartRange)
	catRanges := make(map[int]ChartRange)

	for _, r := range dep.Ranges {
		switch r.Kind {
		case RangeValues:
			if _, ok := valueRanges[r.SeriesIndex]; ok {
				return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("duplicate values range for series %d", r.SeriesIndex)
			}
			valueRanges[r.SeriesIndex] = r
		case RangeCategories:
			if _, ok := catRanges[r.SeriesIndex]; ok {
				return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("duplicate categories range for series %d", r.SeriesIndex)
			}
			catRanges[r.SeriesIndex] = r
		}
	}

	if len(valueRanges) == 0 || len(catRanges) == 0 {
		return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("area chart requires categories and values")
	}

	if len(valueRanges) != len(catRanges) {
		return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("area chart categories/values series mismatch")
	}

	var catKey string
	for idx, cat := range catRanges {
		if _, ok := valueRanges[idx]; !ok {
			return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("area chart categories/values series mismatch")
		}
		key := cat.Sheet + "!" + cat.StartCell + ":" + cat.EndCell
		if catKey == "" {
			catKey = key
		} else if key != catKey {
			return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("area chart categories must match across series")
		}
	}

	return "", nil
}

func (d *Document) handleChartDataMismatch(chartIndex, categoriesLen, valuesLen, seriesIndex int) error {
	if d.opts.Mode != BestEffort {
		return fmt.Errorf("categories length %d does not match values length %d for series %d", categoriesLen, valuesLen, seriesIndex)
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    "CHART_DATA_LENGTH_MISMATCH",
		Message: "Categories and values length mismatch; chart skipped",
		Context: map[string]string{
			"chartIndex":    strconv.Itoa(chartIndex),
			"categoriesLen": strconv.Itoa(categoriesLen),
			"valuesLen":     strconv.Itoa(valuesLen),
			"seriesIndex":   strconv.Itoa(seriesIndex),
		},
	})

	return nil
}

func expandRangeCells(startCell, endCell string) ([]string, error) {
	startCol, startRow, startRef, err := xlref.SplitCellRef(startCell)
	if err != nil {
		return nil, fmt.Errorf("invalid start cell %q: %w", startCell, err)
	}
	endCol, endRow, endRef, err := xlref.SplitCellRef(endCell)
	if err != nil {
		return nil, fmt.Errorf("invalid end cell %q: %w", endCell, err)
	}

	if startCol != endCol && startRow != endRow {
		return nil, fmt.Errorf("2D range %s:%s not supported", startRef, endRef)
	}

	if startCol == endCol {
		if startRow > endRow {
			startRow, endRow = endRow, startRow
		}
		out := make([]string, 0, endRow-startRow+1)
		for row := startRow; row <= endRow; row++ {
			out = append(out, fmt.Sprintf("%s%d", startCol, row))
		}
		return out, nil
	}

	startIdx := colToIndex(startCol)
	endIdx := colToIndex(endCol)
	if startIdx > endIdx {
		startIdx, endIdx = endIdx, startIdx
	}
	out := make([]string, 0, endIdx-startIdx+1)
	for col := startIdx; col <= endIdx; col++ {
		out = append(out, fmt.Sprintf("%s%d", indexToCol(col), startRow))
	}
	return out, nil
}

func colToIndex(col string) int {
	index := 0
	for i := 0; i < len(col); i++ {
		ch := col[i]
		if ch < 'A' || ch > 'Z' {
			return 0
		}
		index = index*26 + int(ch-'A'+1)
	}
	return index
}

func indexToCol(index int) string {
	if index <= 0 {
		return ""
	}
	var buf [8]byte
	pos := len(buf)
	for index > 0 {
		index--
		buf[pos-1] = byte('A' + (index % 26))
		pos--
		index /= 26
	}
	return string(buf[pos:])
}

type noopLogger struct{}

func (noopLogger) Debug(msg string, kv ...any) {}
func (noopLogger) Info(msg string, kv ...any)  {}
func (noopLogger) Warn(msg string, kv ...any)  {}
func (noopLogger) Error(msg string, kv ...any) {}
