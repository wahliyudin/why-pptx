package xlref

import "testing"

func TestParseA1Range(t *testing.T) {
	tests := []struct {
		formula  string
		sheet    string
		start    string
		end      string
		hasError bool
	}{
		{formula: "Sheet1!$A$2:$A$6", sheet: "Sheet1", start: "A2", end: "A6"},
		{formula: "'My Sheet'!B3:D3", sheet: "My Sheet", start: "B3", end: "D3"},
		{formula: "'Data 📈'!$A$2:$A$6", sheet: "Data 📈", start: "A2", end: "A6"},
		{formula: "Sheet1!A2", sheet: "Sheet1", start: "A2", end: "A2"},
		{formula: "Sheet1A2", hasError: true},
		{formula: "Sheet1!", hasError: true},
		{formula: "Sheet1!A0", hasError: true},
	}

	for _, test := range tests {
		ref, err := ParseA1Range(test.formula)
		if test.hasError {
			if err == nil {
				t.Fatalf("expected error for %q", test.formula)
			}
			continue
		}

		if err != nil {
			t.Fatalf("ParseA1Range(%q): %v", test.formula, err)
		}
		if ref.Sheet != test.sheet || ref.StartCell != test.start || ref.EndCell != test.end {
			t.Fatalf("unexpected ref for %q: %+v", test.formula, ref)
		}
	}
}
