# Architecture Notes

## High-level flow

PPTX (zip) -> discover charts -> edit workbook -> sync caches -> save

Discovery finds embedded charts and their workbooks, workbook edits update the
embedded XLSX, cache sync refreshes chart XML caches, and save writes a new PPTX.

## Why ZIP overlay is used

- Avoids extracting to disk.
- Preserves unknown parts for pass-through fidelity.
- Supports atomic save by writing a new ZIP and renaming.

## Why embedded workbooks only

Linked workbooks are external dependencies and are skipped for safety and
reproducibility. Alerts are emitted when linked targets are detected.

## Inline strings only

Strings are written as inlineStr to avoid sharedStrings.xml bookkeeping. This
keeps v0.3 small and avoids cross-part side effects.

## Strict vs BestEffort

- Strict: fail-fast on errors to avoid partial or ambiguous edits.
- BestEffort: skip failing charts/workbooks and emit alerts so batch flows can
  continue with visible diagnostics.

## Cache sync

Charts store cached values in chart XML (c:strCache, c:numCache). After workbook
edits, SyncChartCaches refreshes these caches so PowerPoint displays the new
values immediately.

## Streaming XML transforms

Edits are copy-through transforms to preserve unknown elements, attributes, and
structure. This minimizes unintended diffs and improves compatibility.

## Known limitations (v0.3)

- Bar/line charts only.
- Inline strings only (no sharedStrings).
- 1D ranges only (no 2D ranges).
- No formula evaluation.
