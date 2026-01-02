# Changelog

## v2.0.0

### Added
- Mixed bar+line chart write support, including single-axis and secondary-axis shapes within documented constraints.
- Read-only extraction/export support for pie, area, and mixed bar+line charts.
- Structural snapshot guardrails and alert code baseline checks for release stability.

### Guarantees
- No silent corruption of PPTX/XLSX; per-chart staging with postflight validation before commit.
- Structural determinism: same parts present, relationships resolve, and chart caches match workbook edits when cache sync is enabled.
- Strict vs BestEffort semantics are stable and documented.

### Supported WRITE Scope
- Charts: bar, line, pie (single-series), area (single- or multi-series), mixed bar+line (one bar plot + one line plot).
- Ranges: 1D only.
- Strings: inlineStr only.
- Cache handling: rebuild for referenced ranges; validation enforces cache invariants.

### Known Limitations
- No support for stacked/percent-stacked, 3D, or combined charts beyond bar+line.
- Mixed charts must match documented shapes; secondary axis supported only for bar+line mixed charts with two axis groups.
- No sharedStrings, formula evaluation, 2D ranges, or automatic repair of malformed inputs.
- Output ZIP is not byte-identical to input.

## v1.0.0

### Added
- First stable release of the PPTX manipulation library.
- Staging overlay with postflight validation to prevent partial or unsafe writes.
- Structural regression corpus and snapshot tests.

### Guarantees
- No silent corruption of PPTX/XLSX; postflight validation runs before commit.
- Structural determinism: same parts present, relationships resolve, and chart caches match workbook edits when cache sync is enabled.
- Atomic save via temp file + rename.
- Strict vs BestEffort semantics are stable and documented.
- Alerts are emitted for every skipped or invalid unit of work with stable codes.

### Supported Scope
- Charts: bar, line
- Ranges: 1D only
- Strings: inlineStr only
- Cache handling: rebuild for referenced ranges; validation enforces cache invariants

### Known Limitations
- No support for chart types beyond bar/line.
- No sharedStrings, formula evaluation, 2D ranges, or combined charts.
- No automatic repair of malformed input files.
- Output ZIP is not byte-identical to input.

### Notes
- Determinism is structural, not byte-identical.
- Linked workbooks are detected and skipped with alerts.

## v0.3.2 - Chart Selection DX

### Added
- ListCharts(): list embedded charts with metadata (title, alt text, series count).
- ApplyChartDataByName(): select chart by title or alt text.
- ApplyChartDataByPath(): select chart by chart path.

### Improved
- Developer experience when working with multiple charts.
- Chart discoverability without relying on index order.

### Compatibility
- Fully backward compatible with v0.3.1.
- Existing ApplyChartData(index, ...) unchanged.

### Known limitations
- Bar/line charts only.
- Inline strings only (no sharedStrings).
- 1D ranges only.
- No formula evaluation.
