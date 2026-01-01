# v1.0.0 — Stable, correctness-first PPTX edits

This library edits PPTX files by updating embedded Excel workbooks and synchronizing
chart caches for bar and line charts. It does not attempt to cover the full PPTX
or Excel feature set.

Why v1.0.0 is stable:
- Writes are staged and validated before commit to prevent silent corruption.
- Output is structurally deterministic: parts and relationships remain consistent,
  and caches match workbook edits when cache sync is enabled.
- Behavior is governed by a formal contract and regression corpus.

Strict vs BestEffort:
- Strict: any validation failure aborts the operation.
- BestEffort: a failed chart is skipped, alerts are emitted, and other charts proceed.
  No partial updates are committed.

Alerts:
- Alerts are part of the public contract, with stable codes and structured context.
- Alerts are emitted for every skipped or invalid unit of work and should be handled
  programmatically.

Who this is for:
- Production users who value correctness, repeatability, and clear failure reporting.

Who this is not for:
- Users needing advanced chart types, sharedStrings support, formula evaluation,
  2D ranges, or auto-repair of malformed files.
