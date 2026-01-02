# v2 Release Checklist

- [ ] CI green: go test ./...
- [ ] Corpus regression tests pass (Strict + BestEffort)
- [ ] Golden structural snapshots are up to date (use `-update-golden` intentionally)
- [ ] Determinism double-run guard passes
- [ ] Alert code baseline diff passes (no removals/renames)
- [ ] Entry set unchanged for write tests
- [ ] CONTRACT / ARCHITECTURE / ALERTS docs up to date
- [ ] CHANGELOG updated
- [ ] Tag created for v2.x release
