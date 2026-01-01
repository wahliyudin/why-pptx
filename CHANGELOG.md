# Changelog

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
