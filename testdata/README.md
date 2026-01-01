# PPTX Fixture Corpus

Fixtures under `testdata/pptx/` are minimal PPTX ZIPs used for structural regression tests.

- `bar_simple_embedded.pptx`: Single slide with a bar chart and one series; embedded workbook with categories and values.
- `line_multi_series_embedded.pptx`: Single slide with a line chart and two series; embedded workbook with shared categories and per-series values.
- `linked_workbook_chart.pptx`: Chart points to an external workbook via `TargetMode="External"`; should be skipped with an alert.
- `malformed_chart_cache.pptx`: Chart cache has invalid ptCount/pt entries; postflight cache validation should fail.
- `shared_workbook_two_charts.pptx`: Two charts share one embedded workbook; used to verify per-chart staging and partial success.
- `xlsx_sharedStrings_present.pptx`: Embedded workbook contains `xl/sharedStrings.xml` and a `t="s"` cell; should fail postflight validation.
