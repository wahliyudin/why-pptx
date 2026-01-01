package pptx

import (
	"bytes"
	"fmt"
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
	if dep.ChartType == "mixed" || (dep.ChartType != "bar" && dep.ChartType != "line" && dep.ChartType != "pie" && dep.ChartType != "area") {
		return d.handleChartTypeUnsupported(dep)
	}
	if dep.ChartType == "pie" {
		if code, err := validatePieDependencies(dep); err != nil {
			return d.handlePieWriteError(dep, code, err)
		}
	}
	if dep.ChartType == "area" {
		if code, err := validateAreaDependencies(dep); err != nil {
			return d.handleAreaWriteError(dep, code, err)
		}
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
	if code == "WRITE_AREA_MULTIPLE_SERIES_UNSUPPORTED" {
		message = "Area charts with multiple series are unsupported; chart is skipped"
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
		return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("area chart requires categories and values")
	}
	if len(valueSeries) > 1 || len(catSeries) > 1 {
		return "WRITE_AREA_MULTIPLE_SERIES_UNSUPPORTED", fmt.Errorf("area chart requires exactly one series")
	}

	seriesIndex := -1
	for idx := range valueSeries {
		seriesIndex = idx
	}
	for idx := range catSeries {
		if idx != seriesIndex {
			return "CHART_DEPENDENCIES_PARSE_FAILED", fmt.Errorf("area chart categories/values series mismatch")
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
