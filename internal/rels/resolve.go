package rels

import (
	"path"
	"strings"
)

// ResolveTarget resolves a relationship target against the base part path.
// External targets (TargetMode="External") should be handled by the caller.
func ResolveTarget(basePart string, relTarget string) string {
	if relTarget == "" {
		return ""
	}

	cleanTarget := strings.TrimLeft(relTarget, "/")
	joined := path.Join(path.Dir(basePart), cleanTarget)
	joined = path.Clean(joined)
	return strings.TrimLeft(joined, "/")
}
