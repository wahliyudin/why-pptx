# v1.x Stability & Maintenance Policy

## v1.x Stability Rules (Maintainer Policy)

Allowed in v1.x:
- Bug fixes within documented scope that do not change public API signatures.
- Performance improvements with no semantic change.
- DX polish (logging clarity, error messages, plan/export ergonomics).
- Additional tests (unit, integration, corpus).
- Documentation fixes and clarifications.

Forbidden in v1.x:
- New chart types or range shapes.
- sharedStrings support or formula evaluation.
- Silent behavior changes or weakened validation.
- Changes to Strict vs BestEffort semantics.
- Removing or renaming alert codes.
- Changes that invalidate Step 4 corpus expectations.

## CI Enforcement (Hard Guardrails)

Every v1.x PR must pass:
- Corpus regression tests (Strict + BestEffort).
- Structural determinism checks.
- Exporter determinism tests.
- No unexpected ZIP entry additions/removals.
- Alert code list diff (no removals/renames).
- Plan mode determinism tests.

Recommended:
- Fail CI if CONTRACT.md or ALERTS.md changes without maintainer approval.

## Bugfix Workflow (Safe Flow)

1) Bug report must include:
- Reproduction steps.
- Minimal PPTX fixture if possible.

2) Classification:
- Violates v1 guarantee -> HIGH priority.
- Inside scope but edge-case -> NORMAL.
- Outside scope -> label "v2-candidate".

3) Fix requirements:
- Add a regression test (prefer corpus fixture).
- Document why invariants still hold.

4) Review checklist:
- No scope expansion.
- No validator weakening.
- No silent behavior change.

## Maintainer Review Checklist

Before merging any v1.x PR:
- Does this touch overlaystage, postflight validators, or extraction/export logic?
- If yes, are invariants explicitly discussed and tested?
- Does it affect determinism or ordering?
- Does it affect alerts or mode behavior?
- Is the change clearly safe for v1.x?

## Documentation Hygiene

- CHANGELOG updated for every release.
- ALERTS.md updated if new alert codes are added.
- CONTRACT.md is authoritative; changes require explicit review.
- README and examples must not imply unsupported features.

## Maintenance Mode Statement (User-Facing)

v1.x is in stability and correctness mode. Scope expansion will happen in v2 only.
Bug reports with reproducible fixtures are welcome. Behavior will not change
silently, and compatibility within v1.x is maintained by contract and CI
guardrails.
