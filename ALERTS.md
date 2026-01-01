# Alert Codes

Alert codes are part of the public contract. Codes are emitted in BestEffort
flows; Strict mode generally returns errors instead of recording alerts.

## Discovery and relationships

- CHART_LINKED_WORKBOOK: chart uses a linked workbook and is skipped.
  Context: slide, chart, target
- CHART_RELS_MISSING: chart relationships part is missing; chart is skipped.
  Context: slide, chart, rels_path
- CHART_WORKBOOK_NOT_FOUND: no workbook relationship found for chart.
  Context: slide, chart
- CHART_WORKBOOK_UNSUPPORTED_TARGET: chart target is unsupported.
  Context: slide, chart, target

## Chart parsing and planning

- CHART_DEPENDENCIES_PARSE_FAILED: chart dependencies could not be parsed.
  Context: slide, chart, workbook, error
- CHART_TYPE_UNSUPPORTED: chart type is outside supported scope.
  Context: slide, chart, chartType
- CHART_INFO_PARSE_FAILED: chart metadata parsing failed (ListCharts/Plan).
  Context: slide, chart, error
- CHART_NAME_AMBIGUOUS: chart selection by name is ambiguous.
  Context: name, matches
- CHART_DATA_LENGTH_MISMATCH: categories/values length mismatch.
  Context: chartIndex, categoriesLen, valuesLen, seriesIndex

## Workbook updates

- WORKBOOK_UPDATE_FAILED: workbook cell update failed; workbook is skipped.
  Context: workbook, sheet, cell, error

## Write support

- WRITE_PIE_MULTIPLE_SERIES_UNSUPPORTED: pie chart has multiple series; chart is skipped.
  Context: slide, chart, workbook, error
- WRITE_AREA_MULTIPLE_SERIES_UNSUPPORTED: legacy (pre-multi-series); retained for compatibility.
  Context: slide, chart, workbook, error
- WRITE_AREA_UNSUPPORTED_VARIANT: area chart variant is unsupported; chart is skipped.
  Context: slide, chart, workbook, error

## Cache sync

- CHART_CACHE_SYNC_FAILED: chart cache sync failed; chart is skipped.
  Context: slide, chart, workbook, error

## Postflight validation

- POSTFLIGHT_UNEXPECTED_PART_ADDED: staged update introduced a new part.
  Context: partPath, chartPath, slidePath, workbookPath, stage, mode
- POSTFLIGHT_XML_MALFORMED: malformed XML detected in a touched part.
  Context: partPath, chartPath, slidePath, workbookPath, stage, mode
- POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED: sharedStrings.xml detected in XLSX.
  Context: partPath, workbookPath, stage, mode
- POSTFLIGHT_XLSX_CELL_TYPE_MISMATCH: worksheet contains shared-string cells (t="s").
  Context: workbookPath, sheetPath, cellRef, stage, mode
- POSTFLIGHT_REL_TARGET_MISSING: relationship target missing in package.
  Context: relPath, target, chartPath, slidePath, workbookPath, stage, mode
- POSTFLIGHT_CHART_CACHE_INVALID: chart cache invariants failed.
  Context: chartPath, partPath, seriesIndex, stage, mode

## Read-only extraction

- EXTRACT_INVALID_RANGE: extracted range is invalid or unsupported.
  Context: slide, chart, workbook
- EXTRACT_MIXED_CHART_DETECTED: mixed chart type is unsupported; chart is skipped.
  Context: slide, chart, workbook, error
- EXTRACT_SHAREDSTRINGS_UNSUPPORTED: sharedStrings usage detected in workbook.
  Context: slide, chart, workbook, sheetPath, cell
- EXTRACT_SHEET_NOT_FOUND: referenced sheet name not found in workbook.
  Context: slide, chart, workbook, sheet
- EXTRACT_CELL_PARSE_ERROR: cell value parse failed during extraction/export.
  Context: slide, chart, workbook, sheet, error
- EXPORT_FORMAT_UNSUPPORTED: export format is not registered.
  Context: format
