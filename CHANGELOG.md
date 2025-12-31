# Changelog

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
