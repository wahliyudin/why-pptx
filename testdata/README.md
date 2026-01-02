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
- `mix_bar_line_simple.pptx`: Mixed bar+line chart with shared categories; embedded workbook.
- `mix_bar_line_secondary_axis.pptx`: Mixed bar+line chart with secondary axis IDs; used to validate axis detection.
- `mix_unsupported_variant.pptx`: Mixed chart including an unsupported plot type; extraction should skip or error.
- `mix_write_bar_line_valid.pptx`: Mixed bar+line chart with shared axes; used for write-path edits.
- `mix_write_secondary_axis.pptx`: Mixed bar+line chart with secondary axis IDs and missing axis definitions; used for write-path rejection.
- `mix_write_mismatched_categories.pptx`: Mixed bar+line chart with mismatched category ranges; used for write-path rejection.
- `mix_write_cache_invalid.pptx`: Mixed bar+line chart with invalid cache (ptCount mismatch); used for postflight rejection.
- `mix_write_secondary_axis_valid.pptx`: Mixed bar+line chart with valid secondary axis groups; used for write-path edits.
- `mix_write_secondary_axis_invalid_axis_group.pptx`: Secondary-axis mix with invalid axis group; used for postflight rejection.
- `mix_write_secondary_axis_mismatched_categories.pptx`: Secondary-axis mix with mismatched categories; used for write-path rejection.
- `mix_write_secondary_axis_cache_invalid.pptx`: Secondary-axis mix with invalid cache; used for postflight rejection.
