package chartcache

import (
	"bytes"
	"encoding/xml"
	"io"
	"testing"
)

func TestSyncCachesUpdatesSeriesCaches(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>OldA</c:v></c:pt><c:pt idx="1"><c:v>OldB</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:cat><c:strRef><c:f>Sheet1!$A$4:$A$5</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>OldC</c:v></c:pt><c:pt idx="1"><c:v>OldD</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$4:$B$5</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	deps := Dependencies{
		ChartType: "bar",
		Ranges: []Range{
			{Kind: KindCategories, SeriesIndex: 0, Sheet: "Sheet1", StartCell: "A2", EndCell: "A3"},
			{Kind: KindValues, SeriesIndex: 0, Sheet: "Sheet1", StartCell: "B2", EndCell: "B3"},
			{Kind: KindCategories, SeriesIndex: 1, Sheet: "Sheet1", StartCell: "A4", EndCell: "A5"},
			{Kind: KindValues, SeriesIndex: 1, Sheet: "Sheet1", StartCell: "B4", EndCell: "B5"},
		},
	}

	provider := func(sheet, start, end string) ([]string, error) {
		key := sheet + "!" + start + ":" + end
		switch key {
		case "Sheet1!A2:A3":
			return []string{"Cat1", "Cat2"}, nil
		case "Sheet1!B2:B3":
			return []string{"10", "20"}, nil
		case "Sheet1!A4:A5":
			return []string{"Cat3", "Cat4"}, nil
		case "Sheet1!B4:B5":
			return []string{"30", "40"}, nil
		default:
			return []string{}, nil
		}
	}

	out, err := SyncCaches([]byte(xml), deps, provider)
	if err != nil {
		t.Fatalf("SyncCaches: %v", err)
	}

	cats, nums := extractCacheValues(t, out)
	if len(cats) < 4 || len(nums) < 4 {
		t.Fatalf("expected cache values, got cats=%v nums=%v", cats, nums)
	}
	if cats[0] != "Cat1" || cats[3] != "Cat4" {
		t.Fatalf("unexpected category values: %v", cats)
	}
	if nums[0] != "10" || nums[3] != "40" {
		t.Fatalf("unexpected numeric values: %v", nums)
	}
}

func TestSyncCachesMissingRefErrors(t *testing.T) {
	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat></c:cat>
          <c:val></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

	deps := Dependencies{
		ChartType: "bar",
		Ranges: []Range{
			{Kind: KindCategories, SeriesIndex: 0, Sheet: "Sheet1", StartCell: "A1", EndCell: "A2"},
		},
	}

	_, err := SyncCaches([]byte(xml), deps, func(_, _, _ string) ([]string, error) {
		return []string{"A", "B"}, nil
	})
	if err == nil {
		t.Fatalf("expected error for missing ref nodes")
	}
}

func extractCacheValues(t *testing.T, data []byte) ([]string, []string) {
	t.Helper()

	decoder := xml.NewDecoder(bytes.NewReader(data))
	inStrCache := false
	inNumCache := false
	inValue := false
	var value bytes.Buffer
	var cats []string
	var nums []string

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("decode: %v", err)
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "strCache":
				inStrCache = true
			case "numCache":
				inNumCache = true
			case "v":
				if inStrCache || inNumCache {
					inValue = true
					value.Reset()
				}
			}
		case xml.EndElement:
			switch tok.Name.Local {
			case "strCache":
				inStrCache = false
			case "numCache":
				inNumCache = false
			case "v":
				if inValue {
					if inStrCache {
						cats = append(cats, value.String())
					} else if inNumCache {
						nums = append(nums, value.String())
					}
				}
				inValue = false
			}
		case xml.CharData:
			if inValue {
				value.Write([]byte(tok))
			}
		}
	}

	return cats, nums
}
