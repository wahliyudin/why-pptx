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

## Deprecated options

- `WithBestEffort` is deprecated. Prefer `WithOptions` or `WithErrorMode`.

## Limitations (v0.3)

- Bar/line charts only.
- Inline strings only (no sharedStrings).
- 1D ranges only (no 2D ranges).
- No formula evaluation.
