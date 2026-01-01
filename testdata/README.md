# PPTX Fixture Corpus

Fixtures under `testdata/pptx/` are minimal PPTX ZIPs used for structural regression tests.

- `bar_simple_embedded.pptx`: Single slide with a bar chart and one series; embedded workbook with categories and values.
- `line_multi_series_embedded.pptx`: Single slide with a line chart and two series; embedded workbook with shared categories and per-series values.
- `linked_workbook_chart.pptx`: Chart points to an external workbook via `TargetMode="External"`; should be skipped with an alert.
- `pie_simple_embedded.pptx`: Single slide with a pie chart and one series; embedded workbook with categories and values.
- `area_simple_embedded.pptx`: Single slide with an area chart and one series; embedded workbook with categories and values.
- `pie_linked_workbook.pptx`: Pie chart points to an external workbook; should be skipped with an alert.
- `pie_edit_valid.pptx`: Single-series pie chart with embedded workbook; used for write-path edits.
- `pie_edit_multiple_series.pptx`: Pie chart with multiple series; used to validate write-path rejection.
- `pie_edit_linked_workbook.pptx`: Pie chart with linked workbook; must be skipped with an alert.
- `pie_edit_cache_invalid.pptx`: Pie chart with invalid cache (ptCount/idx); used for postflight rejection.
- `area_edit_valid.pptx`: Single-series area chart with embedded workbook; used for write-path edits.
- `area_edit_multiple_series.pptx`: Area chart with multiple series; legacy write-path fixture.
- `area_edit_linked_workbook.pptx`: Area chart with linked workbook; must be skipped with an alert.
- `area_edit_cache_invalid.pptx`: Area chart with invalid cache (ptCount/idx); used for postflight rejection.
- `area_multi_series_valid.pptx`: Multi-series area chart with shared categories; embedded workbook.
- `area_multi_series_mismatched_categories.pptx`: Area chart with mismatched category ranges; used to validate write-path rejection.
- `area_multi_series_linked_workbook.pptx`: Multi-series area chart with linked workbook; must be skipped with an alert.
- `area_multi_series_cache_invalid.pptx`: Multi-series area chart with invalid cache; used for postflight rejection.
- `malformed_chart_cache.pptx`: Chart cache has invalid ptCount/pt entries; postflight cache validation should fail.
- `shared_workbook_two_charts.pptx`: Two charts share one embedded workbook; used to verify per-chart staging and partial success.
- `xlsx_sharedStrings_present.pptx`: Embedded workbook contains `xl/sharedStrings.xml` and a `t="s"` cell; should fail postflight validation.
