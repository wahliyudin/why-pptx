package pptx

import "testing"

func TestAlertsDefensiveCopy(t *testing.T) {
	doc := &Document{}
	doc.addAlert(Alert{
		Level:   "warn",
		Code:    "W001",
		Message: "test",
		Context: map[string]string{"key": "value"},
	})

	alerts := doc.Alerts()
	if len(alerts) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(alerts))
	}

	alerts[0].Level = "changed"
	alerts[0].Context["key"] = "changed"

	if doc.alerts[0].Level != "warn" {
		t.Fatalf("internal alert mutated: %q", doc.alerts[0].Level)
	}
	if doc.alerts[0].Context["key"] != "value" {
		t.Fatalf("internal context mutated: %q", doc.alerts[0].Context["key"])
	}
}

func TestAlertsEmptySlice(t *testing.T) {
	var nilDoc *Document
	if alerts := nilDoc.Alerts(); alerts == nil || len(alerts) != 0 {
		t.Fatalf("expected empty slice for nil doc, got %#v", alerts)
	}
	for range nilDoc.Alerts() {
		t.Fatal("unexpected alert in nil doc")
	}

	doc := &Document{}
	if alerts := doc.Alerts(); alerts == nil || len(alerts) != 0 {
		t.Fatalf("expected empty slice, got %#v", alerts)
	}
	for range doc.Alerts() {
		t.Fatal("unexpected alert in empty doc")
	}
}

func TestDocumentCloseNoop(t *testing.T) {
	doc := &Document{}
	if err := doc.Close(); err != nil {
		t.Fatalf("Close error: %v", err)
	}
}

func TestAlertHelpers(t *testing.T) {
	doc := &Document{}
	if doc.HasAlerts() {
		t.Fatalf("expected no alerts")
	}

	doc.addAlert(Alert{Level: "warn", Code: "W001", Message: "first", Context: map[string]string{"k": "v"}})
	doc.addAlert(Alert{Level: "info", Code: "I002", Message: "second"})

	if !doc.HasAlerts() {
		t.Fatalf("expected alerts")
	}

	matched := doc.AlertsByCode("W001")
	if len(matched) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(matched))
	}
	if matched[0].Code != "W001" {
		t.Fatalf("unexpected alert code: %q", matched[0].Code)
	}
	matched[0].Level = "changed"
	matched[0].Context["k"] = "changed"
	if doc.alerts[0].Level != "warn" {
		t.Fatalf("internal alert mutated: %q", doc.alerts[0].Level)
	}
	if doc.alerts[0].Context["k"] != "v" {
		t.Fatalf("internal context mutated: %q", doc.alerts[0].Context["k"])
	}

	none := doc.AlertsByCode("missing")
	if none == nil || len(none) != 0 {
		t.Fatalf("expected empty slice, got %#v", none)
	}
}
