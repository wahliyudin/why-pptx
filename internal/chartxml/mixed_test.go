package chartxml

import (
	"strings"
	"testing"
)

func TestParseMixedBarLine(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:lineChart>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	mixed, err := ParseMixed(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("ParseMixed: %v", err)
	}
	if len(mixed.Series) != 2 {
		t.Fatalf("expected 2 series, got %d", len(mixed.Series))
	}
	if mixed.Series[0].PlotType != "bar" || mixed.Series[1].PlotType != "line" {
		t.Fatalf("unexpected plot types: %+v", mixed.Series)
	}
	if mixed.Series[0].Axis != "primary" || mixed.Series[1].Axis != "primary" {
		t.Fatalf("expected primary axes, got %+v", mixed.Series)
	}
	if len(mixed.Series[0].Formulas) == 0 || len(mixed.Series[1].Formulas) == 0 {
		t.Fatalf("expected formulas in mixed series")
	}
}

func TestParseMixedSecondaryAxis(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat></c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:lineChart>
        <c:ser><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat></c:ser>
        <c:axId val="3"/>
        <c:axId val="4"/>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	mixed, err := ParseMixed(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("ParseMixed: %v", err)
	}
	if len(mixed.Series) != 2 {
		t.Fatalf("expected 2 series, got %d", len(mixed.Series))
	}
	if mixed.Series[0].Axis != "primary" || mixed.Series[1].Axis != "secondary" {
		t.Fatalf("expected secondary axis on second series, got %+v", mixed.Series)
	}
}

func TestParseMixedUnsupportedPlot(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart><c:ser></c:ser></c:barChart>
      <c:areaChart><c:ser></c:ser></c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	if _, err := ParseMixed(strings.NewReader(xml)); err == nil {
		t.Fatalf("expected ParseMixed error for unsupported plot")
	}
}

func TestParseMixedStackedUnsupported(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:grouping val="stacked"/>
        <c:ser></c:ser>
      </c:barChart>
      <c:lineChart><c:ser></c:ser></c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	if _, err := ParseMixed(strings.NewReader(xml)); err == nil {
		t.Fatalf("expected ParseMixed error for stacked chart")
	}
}
