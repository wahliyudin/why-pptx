# Architecture Notes

## Cache sync

PPTX charts store cached values inside the chart XML (`c:strCache`, `c:numCache`).
After editing embedded workbooks, `SyncChartCaches` refreshes these caches so
PowerPoint renders updated values immediately.

## Inline strings only

V0 writes strings as `inlineStr` to avoid shared string tables. This keeps the
scope small and avoids cross-part bookkeeping. Shared strings may be added later.

## Streaming XML transforms

XML edits are done as copy-through streaming transforms to preserve unknown
elements, attributes, and overall structure. This minimizes unintended diffs and
improves compatibility with existing PPTX/XLSX features.
