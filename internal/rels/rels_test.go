package rels

import (
	"strings"
	"testing"
)

func TestParseRelsInternal(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`

	rels, err := Parse(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	rel, ok := rels.Resolve("rId1")
	if !ok {
		t.Fatalf("expected relationship rId1")
	}
	if rel.Target != "../charts/chart1.xml" {
		t.Fatalf("unexpected target: %q", rel.Target)
	}
}

func TestParseRelsExternal(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="https://example.com/book.xlsx" TargetMode="External"/>
</Relationships>`

	rels, err := Parse(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}

	rel, ok := rels.Resolve("rId2")
	if !ok {
		t.Fatalf("expected relationship rId2")
	}
	if rel.TargetMode != "External" {
		t.Fatalf("unexpected target mode: %q", rel.TargetMode)
	}
}

func TestResolveTarget(t *testing.T) {
	got := ResolveTarget("ppt/slides/slide1.xml", "../charts/chart1.xml")
	if got != "ppt/charts/chart1.xml" {
		t.Fatalf("unexpected resolved path: %q", got)
	}
}
