package pptx

import (
	"fmt"

	"why-pptx/internal/chartdiscover"
	"why-pptx/internal/ooxmlpkg"
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
	pkg    *ooxmlpkg.Package
	alerts []Alert
	logger Logger
	strict bool
}

type EmbeddedChart struct {
	SlidePath    string
	ChartPath    string
	WorkbookPath string
}

func OpenFile(path string, opts ...Option) (*Document, error) {
	pkg, err := ooxmlpkg.OpenFile(path)
	if err != nil {
		return nil, err
	}

	doc := &Document{
		pkg:    pkg,
		logger: noopLogger{},
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

type noopLogger struct{}

func (noopLogger) Debug(msg string, kv ...any) {}
func (noopLogger) Info(msg string, kv ...any)  {}
func (noopLogger) Warn(msg string, kv ...any)  {}
func (noopLogger) Error(msg string, kv ...any) {}
