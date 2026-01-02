package pptx

import (
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"
)

func TestAlertCodeBaseline(t *testing.T) {
	baselinePath := filepath.Join("..", "testdata", "alert_codes_baseline.txt")
	baseline, err := readLines(baselinePath)
	if err != nil {
		t.Fatalf("read baseline: %v", err)
	}

	current, err := readAlertCodes(filepath.Join("..", "ALERTS.md"))
	if err != nil {
		t.Fatalf("read ALERTS.md: %v", err)
	}

	var missing []string
	for _, code := range baseline {
		if _, ok := current[code]; !ok {
			missing = append(missing, code)
		}
	}
	sort.Strings(missing)

	if len(missing) > 0 {
		t.Fatalf("alert codes removed or renamed: %v", missing)
	}
}

func readAlertCodes(path string) (map[string]struct{}, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}

	out := make(map[string]struct{})
	for _, line := range strings.Split(string(data), "\n") {
		line = strings.TrimSpace(line)
		if !strings.HasPrefix(line, "- ") {
			continue
		}
		trimmed := strings.TrimPrefix(line, "- ")
		parts := strings.SplitN(trimmed, ":", 2)
		if len(parts) == 0 {
			continue
		}
		code := strings.TrimSpace(parts[0])
		if code == "" {
			continue
		}
		out[code] = struct{}{}
	}
	return out, nil
}

func readLines(path string) ([]string, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, err
	}
	lines := strings.Split(string(data), "\n")
	out := make([]string, 0, len(lines))
	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line == "" {
			continue
		}
		out = append(out, line)
	}
	return out, nil
}
