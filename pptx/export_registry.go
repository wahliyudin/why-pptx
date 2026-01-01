package pptx

import (
	"fmt"
	"sort"
	"sync"
)

// ExporterRegistry stores exporters by format in a thread-safe map.
type ExporterRegistry struct {
	mu        sync.RWMutex
	exporters map[ExportFormat]Exporter
}

func NewExporterRegistry() *ExporterRegistry {
	return &ExporterRegistry{
		exporters: make(map[ExportFormat]Exporter),
	}
}

// DefaultExporterRegistry returns a new registry with built-in exporters.
// It does not share state across calls.
func DefaultExporterRegistry() *ExporterRegistry {
	return defaultExporterRegistry(DefaultOptions())
}

func defaultExporterRegistry(opts Options) *ExporterRegistry {
	reg := NewExporterRegistry()
	_ = reg.Register(ChartJSExporter{MissingNumericPolicy: opts.Workbook.MissingNumericPolicy})
	return reg
}

func (r *ExporterRegistry) Register(exp Exporter) error {
	if r == nil {
		return fmt.Errorf("exporter registry is nil")
	}
	if exp == nil {
		return fmt.Errorf("exporter is nil")
	}
	format := exp.Format()
	if format == "" {
		return fmt.Errorf("exporter format is empty")
	}

	r.mu.Lock()
	defer r.mu.Unlock()

	if _, exists := r.exporters[format]; exists {
		return fmt.Errorf("exporter format %q already registered", format)
	}
	r.exporters[format] = exp
	return nil
}

func (r *ExporterRegistry) MustRegister(exp Exporter) {
	if err := r.Register(exp); err != nil {
		panic(err)
	}
}

func (r *ExporterRegistry) Get(format ExportFormat) (Exporter, bool) {
	if r == nil {
		return nil, false
	}
	r.mu.RLock()
	defer r.mu.RUnlock()
	exp, ok := r.exporters[format]
	return exp, ok
}

func (r *ExporterRegistry) Formats() []ExportFormat {
	if r == nil {
		return []ExportFormat{}
	}
	r.mu.RLock()
	defer r.mu.RUnlock()

	formats := make([]ExportFormat, 0, len(r.exporters))
	for format := range r.exporters {
		formats = append(formats, format)
	}
	sort.Slice(formats, func(i, j int) bool {
		return formats[i] < formats[j]
	})
	return formats
}

func (r *ExporterRegistry) Unregister(format ExportFormat) {
	if r == nil || format == "" {
		return
	}
	r.mu.Lock()
	defer r.mu.Unlock()
	delete(r.exporters, format)
}
