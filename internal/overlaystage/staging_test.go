package overlaystage

import (
	"bytes"
	"sort"
	"testing"
)

type memOverlay struct {
	baseline map[string][]byte
	overlay  map[string][]byte
}

func newMemOverlay(parts map[string][]byte) *memOverlay {
	baseline := make(map[string][]byte, len(parts))
	for name, data := range parts {
		copied := make([]byte, len(data))
		copy(copied, data)
		baseline[name] = copied
	}
	return &memOverlay{
		baseline: baseline,
		overlay:  make(map[string][]byte),
	}
}

func (m *memOverlay) Get(path string) ([]byte, error) {
	if data, ok := m.overlay[path]; ok {
		return append([]byte(nil), data...), nil
	}
	if data, ok := m.baseline[path]; ok {
		return append([]byte(nil), data...), nil
	}
	return nil, errNotFound(path)
}

func (m *memOverlay) Set(path string, content []byte) error {
	copied := make([]byte, len(content))
	copy(copied, content)
	m.overlay[path] = copied
	return nil
}

func (m *memOverlay) Has(path string) (bool, error) {
	if _, ok := m.overlay[path]; ok {
		return true, nil
	}
	_, ok := m.baseline[path]
	return ok, nil
}

func (m *memOverlay) ListEntries() ([]string, error) {
	seen := make(map[string]struct{}, len(m.baseline)+len(m.overlay))
	names := make([]string, 0, len(m.baseline)+len(m.overlay))
	for name := range m.baseline {
		seen[name] = struct{}{}
		names = append(names, name)
	}
	for name := range m.overlay {
		if _, ok := seen[name]; ok {
			continue
		}
		names = append(names, name)
	}
	sort.Strings(names)
	return names, nil
}

func (m *memOverlay) HasBaseline(path string) (bool, error) {
	_, ok := m.baseline[path]
	return ok, nil
}

type notFoundError struct {
	path string
}

func (e notFoundError) Error() string {
	return "not found: " + e.path
}

func errNotFound(path string) error {
	return notFoundError{path: path}
}

func TestStagingOverlayIsolation(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("orig"),
	})
	stage := NewStagingOverlay(parent)

	if err := stage.Set("ppt/charts/chart1.xml", []byte("updated")); err != nil {
		t.Fatalf("Set: %v", err)
	}

	parentData, err := parent.Get("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("parent.Get: %v", err)
	}
	if !bytes.Equal(parentData, []byte("orig")) {
		t.Fatalf("parent data changed: %q", parentData)
	}

	stageData, err := stage.Get("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("stage.Get: %v", err)
	}
	if !bytes.Equal(stageData, []byte("updated")) {
		t.Fatalf("unexpected stage data: %q", stageData)
	}
}

func TestStagingOverlayDiscard(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("orig"),
	})
	stage := NewStagingOverlay(parent)

	if err := stage.Set("ppt/charts/chart1.xml", []byte("updated")); err != nil {
		t.Fatalf("Set: %v", err)
	}
	stage.Discard()

	parentData, err := parent.Get("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("parent.Get: %v", err)
	}
	if !bytes.Equal(parentData, []byte("orig")) {
		t.Fatalf("parent data changed: %q", parentData)
	}
}

func TestStagingOverlayCommit(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("orig"),
	})
	stage := NewStagingOverlay(parent)

	if err := stage.Set("ppt/charts/chart1.xml", []byte("updated")); err != nil {
		t.Fatalf("Set: %v", err)
	}
	if err := stage.Commit(); err != nil {
		t.Fatalf("Commit: %v", err)
	}

	parentData, err := parent.Get("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("parent.Get: %v", err)
	}
	if !bytes.Equal(parentData, []byte("updated")) {
		t.Fatalf("commit not applied: %q", parentData)
	}
}

func TestStagingOverlayListTouchedSorted(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("orig"),
		"ppt/charts/chart2.xml": []byte("orig2"),
	})
	stage := NewStagingOverlay(parent)

	_ = stage.Set("ppt/charts/chart2.xml", []byte("b"))
	_ = stage.Set("ppt/charts/chart1.xml", []byte("a"))

	touched := stage.ListTouched()
	want := []string{"ppt/charts/chart1.xml", "ppt/charts/chart2.xml"}
	if len(touched) != len(want) {
		t.Fatalf("expected %d touched entries, got %d", len(want), len(touched))
	}
	for i := range want {
		if touched[i] != want[i] {
			t.Fatalf("unexpected touched order: %v", touched)
		}
	}
}

func TestStagingOverlayCommitNoNewEntries(t *testing.T) {
	parent := newMemOverlay(map[string][]byte{
		"ppt/charts/chart1.xml": []byte("orig"),
	})
	stage := NewStagingOverlay(parent)

	if err := stage.Set("ppt/charts/new.xml", []byte("data")); err != nil {
		t.Fatalf("Set: %v", err)
	}
	if err := stage.Commit(); err == nil {
		t.Fatalf("expected commit error for new entry")
	}

	parentData, err := parent.Get("ppt/charts/chart1.xml")
	if err != nil {
		t.Fatalf("parent.Get: %v", err)
	}
	if !bytes.Equal(parentData, []byte("orig")) {
		t.Fatalf("parent data changed: %q", parentData)
	}
}
