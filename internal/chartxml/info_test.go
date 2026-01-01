package chartxml

import (
	"strings"
	"testing"
)

func TestParseInfoPieChart(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser></c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	info, err := ParseInfo(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("ParseInfo: %v", err)
	}
	if info.ChartType != "pie" {
		t.Fatalf("expected pie chart type, got %q", info.ChartType)
	}
	if info.SeriesCount != 1 {
		t.Fatalf("expected series count 1, got %d", info.SeriesCount)
	}
}

func TestParseInfoAreaChart(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:ser></c:ser>
        <c:ser></c:ser>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	info, err := ParseInfo(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("ParseInfo: %v", err)
	}
	if info.ChartType != "area" {
		t.Fatalf("expected area chart type, got %q", info.ChartType)
	}
	if info.SeriesCount != 2 {
		t.Fatalf("expected series count 2, got %d", info.SeriesCount)
	}
}

func TestParseInfoMixedChart(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser></c:ser></c:barChart>
      <c:lineChart><c:ser></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	info, err := ParseInfo(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("ParseInfo: %v", err)
	}
	if info.ChartType != "mixed" {
		t.Fatalf("expected mixed chart type, got %q", info.ChartType)
	}
	if info.SeriesCount != 2 {
		t.Fatalf("expected series count 2, got %d", info.SeriesCount)
	}
}
