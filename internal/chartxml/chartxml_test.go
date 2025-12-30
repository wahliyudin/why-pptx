package chartxml

import (
	"strings"
	"testing"

	"why-pptx/internal/xlref"
)

func TestParseBarChartFormulas(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$4:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$4:$B$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	parsed, err := Parse(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if parsed.ChartType != "bar" {
		t.Fatalf("expected bar chart type, got %q", parsed.ChartType)
	}
	if len(parsed.Formulas) != 4 {
		t.Fatalf("expected 4 formulas, got %d", len(parsed.Formulas))
	}

	for _, formula := range parsed.Formulas {
		ref, err := xlref.ParseA1Range(formula.Formula)
		if err != nil {
			t.Fatalf("ParseA1Range(%q): %v", formula.Formula, err)
		}
		if ref.Sheet != "Sheet1" {
			t.Fatalf("unexpected sheet: %q", ref.Sheet)
		}
	}
}

func TestParseLineChartWithSeriesName(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:tx><c:strRef><c:f>'My Sheet'!$B$3:$D$3</c:f></c:strRef></c:tx>
          <c:val><c:numRef><c:f>Sheet1!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	parsed, err := Parse(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if parsed.ChartType != "line" {
		t.Fatalf("expected line chart type, got %q", parsed.ChartType)
	}
	if len(parsed.Formulas) != 2 {
		t.Fatalf("expected 2 formulas, got %d", len(parsed.Formulas))
	}

	foundName := false
	for _, formula := range parsed.Formulas {
		if formula.Kind == KindSeriesName {
			if formula.SeriesIndex != 0 {
				t.Fatalf("unexpected series index: %d", formula.SeriesIndex)
			}
			ref, err := xlref.ParseA1Range(formula.Formula)
			if err != nil {
				t.Fatalf("ParseA1Range(%q): %v", formula.Formula, err)
			}
			if ref.Sheet != "My Sheet" || ref.StartCell != "B3" || ref.EndCell != "D3" {
				t.Fatalf("unexpected range: %+v", ref)
			}
			foundName = true
		}
	}
	if !foundName {
		t.Fatalf("expected series name formula")
	}
}

func TestParseIgnoresEmptyFormulas(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>   </c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f> Sheet1!A1:A2 </c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	parsed, err := Parse(strings.NewReader(xml))
	if err != nil {
		t.Fatalf("Parse: %v", err)
	}
	if len(parsed.Formulas) != 1 {
		t.Fatalf("expected 1 formula, got %d", len(parsed.Formulas))
	}
	if parsed.Formulas[0].Formula != "Sheet1!A1:A2" {
		t.Fatalf("unexpected formula: %q", parsed.Formulas[0].Formula)
	}
}
