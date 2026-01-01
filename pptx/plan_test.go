package pptx

import "testing"

func TestPlanLinkedWorkbookSkip(t *testing.T) {
	doc, err := OpenFile(fixturePath("linked_workbook_chart.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	plan, err := doc.Plan()
	if err != nil {
		t.Fatalf("Plan: %v", err)
	}
	if len(plan.Charts) != 1 {
		t.Fatalf("expected 1 chart, got %d", len(plan.Charts))
	}

	chart := plan.Charts[0]
	if chart.Action != "linked" {
		t.Fatalf("expected action linked, got %q", chart.Action)
	}
	if chart.ReasonCode != "CHART_LINKED_WORKBOOK" {
		t.Fatalf("expected reason CHART_LINKED_WORKBOOK, got %q", chart.ReasonCode)
	}

	if len(plan.Alerts) != 1 || plan.Alerts[0].Code != "CHART_LINKED_WORKBOOK" {
		t.Fatalf("expected CHART_LINKED_WORKBOOK alert, got %#v", plan.Alerts)
	}
	if len(doc.Alerts()) != 0 {
		t.Fatalf("Plan should not mutate document alerts")
	}
}

func TestPlanMalformedChartCacheApply(t *testing.T) {
	doc, err := OpenFile(fixturePath("malformed_chart_cache.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	plan, err := doc.Plan()
	if err != nil {
		t.Fatalf("Plan: %v", err)
	}
	if len(plan.Charts) != 1 {
		t.Fatalf("expected 1 chart, got %d", len(plan.Charts))
	}
	if plan.Charts[0].Action != "apply" {
		t.Fatalf("expected action apply, got %q", plan.Charts[0].Action)
	}
}

func TestPlanOrderStable(t *testing.T) {
	doc, err := OpenFile(fixturePath("shared_workbook_two_charts.pptx"))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	plan, err := doc.Plan()
	if err != nil {
		t.Fatalf("Plan: %v", err)
	}
	if len(plan.Charts) != 2 {
		t.Fatalf("expected 2 charts, got %d", len(plan.Charts))
	}
	if plan.Charts[0].ChartPath != "ppt/charts/chart1.xml" || plan.Charts[0].Index != 0 {
		t.Fatalf("unexpected chart order: %#v", plan.Charts)
	}
	if plan.Charts[1].ChartPath != "ppt/charts/chart2.xml" || plan.Charts[1].Index != 1 {
		t.Fatalf("unexpected chart order: %#v", plan.Charts)
	}
}
