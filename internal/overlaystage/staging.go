package overlaystage

import (
	"fmt"
	"sort"
)

type StagingOverlay struct {
	parent Overlay
	staged map[string][]byte
}

func NewStagingOverlay(parent Overlay) *StagingOverlay {
	return &StagingOverlay{
		parent: parent,
		staged: make(map[string][]byte),
	}
}

func (s *StagingOverlay) Get(path string) ([]byte, error) {
	if s == nil || s.parent == nil {
		return nil, fmt.Errorf("stage not initialized")
	}
	if data, ok := s.staged[path]; ok {
		return append([]byte(nil), data...), nil
	}
	return s.parent.Get(path)
}

func (s *StagingOverlay) Set(path string, content []byte) error {
	if s == nil {
		return fmt.Errorf("stage not initialized")
	}
	copied := make([]byte, len(content))
	copy(copied, content)
	s.staged[path] = copied
	return nil
}

func (s *StagingOverlay) Has(path string) (bool, error) {
	if s == nil || s.parent == nil {
		return false, fmt.Errorf("stage not initialized")
	}
	if _, ok := s.staged[path]; ok {
		return true, nil
	}
	return s.parent.Has(path)
}

func (s *StagingOverlay) ListEntries() ([]string, error) {
	if s == nil || s.parent == nil {
		return nil, fmt.Errorf("stage not initialized")
	}
	return s.parent.ListEntries()
}

func (s *StagingOverlay) ListTouched() []string {
	if s == nil || len(s.staged) == 0 {
		return []string{}
	}
	names := make([]string, 0, len(s.staged))
	for name := range s.staged {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}

func (s *StagingOverlay) Commit() error {
	if s == nil || s.parent == nil {
		return fmt.Errorf("stage not initialized")
	}
	if len(s.staged) == 0 {
		return nil
	}

	paths := s.ListTouched()
	for _, path := range paths {
		exists, err := s.hasBaseline(path)
		if err != nil {
			return fmt.Errorf("check baseline for %q: %w", path, err)
		}
		if !exists {
			return fmt.Errorf("commit staged part %q: part does not exist in baseline", path)
		}
	}

	for _, path := range paths {
		if err := s.parent.Set(path, s.staged[path]); err != nil {
			return fmt.Errorf("commit staged part %q: %w", path, err)
		}
	}

	s.staged = make(map[string][]byte)
	return nil
}

func (s *StagingOverlay) Discard() {
	if s == nil {
		return
	}
	s.staged = make(map[string][]byte)
}

func (s *StagingOverlay) Nested() *StagingOverlay {
	if s == nil {
		return &StagingOverlay{}
	}
	return NewStagingOverlay(s)
}

func (s *StagingOverlay) hasBaseline(path string) (bool, error) {
	if checker, ok := s.parent.(BaselineChecker); ok {
		return checker.HasBaseline(path)
	}
	return s.parent.Has(path)
}
