# why-pptx

Small Go library for manipulating PPTX charts and their embedded workbooks.
V0.3 focuses on bar/line charts, inline strings, and fidelity-preserving ZIP/XML handling.

## Basic usage

```go
opts := pptx.DefaultOptions()
opts.Mode = pptx.BestEffort
opts.Chart.CacheSync = true
opts.Workbook.MissingNumericPolicy = pptx.MissingNumericZero

doc, err := pptx.OpenFile("in.pptx", pptx.WithOptions(opts))
if err != nil {
	// handle error
}
defer doc.Close()

err = doc.SetWorkbookCells([]pptx.CellUpdate{
	{
		WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx",
		Sheet:        "Sheet1",
		Cell:         "B3",
		Value:        pptx.Num(42),
	},
})
if err != nil {
	// handle error
}

if err := doc.SyncChartCaches(); err != nil {
	// handle error
}

if err := doc.SaveFile("out.pptx"); err != nil {
	// handle error
}
```

## ApplyChartData example

```go
opts := pptx.DefaultOptions()
opts.Chart.CacheSync = false // skip cache sync for speed

doc, err := pptx.OpenFile("in.pptx", pptx.WithOptions(opts))
if err != nil {
	// handle error
}
defer doc.Close()

err = doc.ApplyChartData(0, map[string][]string{
	"categories": {"Q1", "Q2"},
	"values:0":   {"10", "20"},
})
if err != nil {
	// handle error
}

if err := doc.SaveFile("out.pptx"); err != nil {
	// handle error
}
```

## List charts by title

```go
charts, err := doc.ListCharts()
if err != nil {
	// handle error
}
for _, chart := range charts {
	// chart.Title, chart.ChartPath, chart.SeriesCount...
}

err = doc.ApplyChartDataByName("Revenue", map[string][]string{
	"categories": {"Q1", "Q2"},
	"values:0":   {"10", "20"},
})
if err != nil {
	// handle error (ambiguous or not found)
}
```

If multiple charts share the same title/alt text, ApplyChartDataByName returns
an error (BestEffort also emits a CHART_NAME_AMBIGUOUS alert).

## Plan mode (dry-run)

PlanChanges computes what would be applied or skipped without modifying the
document or writing output. It uses Strict/BestEffort to classify skips vs
errors and does not run postflight validation, so some issues may only surface
during apply.
Unsupported chart types are marked with Action=unsupported and ReasonCode=CHART_TYPE_UNSUPPORTED.

```go
plan, err := doc.Plan()
if err != nil {
	// handle error
}
for _, chart := range plan.Charts {
	// chart.Action, chart.ReasonCode, chart.Dependencies...
}
```

## Read-only extraction and export

ExtractChartDataByPath reads embedded workbook values without modifying the PPTX.
Single-chart extraction returns an error on unsupported input in both modes.
In BestEffort, use ExtractAllCharts/ExportAllCharts to skip charts with alerts.

```go
doc, err := pptx.OpenFile("in.pptx")
if err != nil {
	// handle error
}
defer doc.Close()

data, err := doc.ExtractChartDataByPath("ppt/charts/chart1.xml")
if err != nil {
	// handle error
}

payload, err := doc.ExportChartByPath("ppt/charts/chart1.xml", pptx.ChartJSExporter{
	MissingNumericPolicy: pptx.MissingNumericEmpty,
})
if err != nil {
	// handle error
}
```

## Options

- `Options.Mode`: `Strict` (default) or `BestEffort`.
- `Options.Chart.CacheSync`: update chart caches after workbook edits (default true).
- `Options.Workbook.MissingNumericPolicy`: `MissingNumericEmpty` (default) or `MissingNumericZero`.

`WithOptions` replaces the full options struct; use `DefaultOptions()` as a base.

## Alerts

Alerts are recorded on `Document` in best-effort flows:

- `Alerts()` returns a defensive copy.
- `HasAlerts()` checks if any alerts were emitted.
- `AlertsByCode(code)` filters by code.

## Convenience API

`ApplyChartData` lets you update categories and series values by chart index.

## Migration (v0.2 -> v0.3)

Before (v0.2):

```go
doc, err := pptx.OpenFile("in.pptx", pptx.WithBestEffort(true))
```

After (v0.3):

```go
opts := pptx.DefaultOptions()
opts.Mode = pptx.BestEffort
doc, err := pptx.OpenFile("in.pptx", pptx.WithOptions(opts))
```

## Deprecated options

- `WithBestEffort` is deprecated. Prefer `WithOptions` or `WithErrorMode`.

## Limitations (v0.3)

- Bar/line charts only.
- Inline strings only (no sharedStrings).
- 1D ranges only (no 2D ranges).
- No formula evaluation.
