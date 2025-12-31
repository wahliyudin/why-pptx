package pptx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"testing"
)

func TestSyncChartCachesIntegration(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	workbook := buildWorkbookWithValues(t, "Cat1", "Cat2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithCaches("Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3", []string{"Old1", "Old2"}, []string{"1", "2"}),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	catA := "New1"
	catB := "New2"
	valA := 100.0
	valB := 200.0
	if err := doc.SetWorkbookCells([]CellUpdate{
		{WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx", Sheet: "Sheet1", Cell: "A2", Value: Str(catA)},
		{WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx", Sheet: "Sheet1", Cell: "A3", Value: Str(catB)},
		{WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx", Sheet: "Sheet1", Cell: "B2", Value: Num(valA)},
		{WorkbookPath: "ppt/embeddings/embeddedWorkbook1.xlsx", Sheet: "Sheet1", Cell: "B3", Value: Num(valB)},
	}); err != nil {
		t.Fatalf("SetWorkbookCells: %v", err)
	}

	if err := doc.SyncChartCaches(); err != nil {
		t.Fatalf("SyncChartCaches: %v", err)
	}
	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	chartXML := readZipEntry(t, outputPath, "ppt/charts/chart1.xml")
	cats, nums := extractChartCacheValues(t, chartXML)
	if len(cats) != 2 || cats[0] != "New1" || cats[1] != "New2" {
		t.Fatalf("unexpected category cache: %v", cats)
	}
	if len(nums) != 2 || nums[0] != "100" || nums[1] != "200" {
		t.Fatalf("unexpected value cache: %v", nums)
	}
}

func TestSyncChartCachesBestEffort(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	workbook := buildWorkbookWithValues(t, "Cat1", "Cat2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithCaches("Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3", []string{"Old1", "Old2"}, []string{"1", "2"}),
		"ppt/charts/chart2.xml": chartWithCaches("Sheet1!$A$1:$B$2", "Sheet1!$B$1:$C$2", []string{"X", "Y"}, []string{"3", "4"}),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/charts/_rels/chart2.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(inputPath, WithErrorMode(BestEffort))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if err := doc.SyncChartCaches(); err != nil {
		t.Fatalf("SyncChartCaches: %v", err)
	}
	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	alerts := doc.Alerts()
	if len(alerts) != 1 {
		t.Fatalf("expected 1 alert, got %d", len(alerts))
	}
	if alerts[0].Code != "CHART_CACHE_SYNC_FAILED" {
		t.Fatalf("unexpected alert code: %q", alerts[0].Code)
	}

	chartXML := readZipEntry(t, outputPath, "ppt/charts/chart1.xml")
	cats, nums := extractChartCacheValues(t, chartXML)
	if len(cats) != 2 || cats[0] != "Cat1" || cats[1] != "Cat2" {
		t.Fatalf("unexpected category cache: %v", cats)
	}
	if len(nums) != 2 || nums[0] != "10" || nums[1] != "20" {
		t.Fatalf("unexpected value cache: %v", nums)
	}
}

func TestSyncChartCachesStrictFails(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")

	workbook := buildWorkbookWithValues(t, "Cat1", "Cat2", 10, 20)
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartWithCaches("Sheet1!$A$1:$B$2", "Sheet1!$B$1:$C$2", []string{"X", "Y"}, []string{"3", "4"}),
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	doc, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if err := doc.SyncChartCaches(); err == nil {
		t.Fatalf("expected error in strict mode")
	}
}

func TestSyncChartCachesDisabledNoop(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	workbook := buildWorkbookWithValues(t, "Cat1", "Cat2", 10, 20)
	chartXML := chartWithCaches("Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3", []string{"Old1", "Old2"}, []string{"1", "2"})
	parts := map[string][]byte{
		"ppt/slides/slide1.xml": []byte("<slide/>"),
		"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
		"ppt/charts/chart1.xml": chartXML,
		"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
		"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
	}

	if err := writeZipFile(inputPath, parts); err != nil {
		t.Fatalf("writeZipFile: %v", err)
	}

	opts := DefaultOptions()
	opts.Chart.CacheSync = false
	doc, err := OpenFile(inputPath, WithOptions(opts))
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if err := doc.SyncChartCaches(); err != nil {
		t.Fatalf("SyncChartCaches: %v", err)
	}
	if err := doc.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	outChart := readZipEntry(t, outputPath, "ppt/charts/chart1.xml")
	if !bytes.Equal(outChart, chartXML) {
		t.Fatalf("chart xml changed with CacheSync disabled")
	}
}

func TestSyncChartCachesMissingNumericPolicy(t *testing.T) {
	cases := []struct {
		name   string
		policy MissingNumericPolicy
		want   string
	}{
		{name: "empty", policy: MissingNumericEmpty, want: ""},
		{name: "zero", policy: MissingNumericZero, want: "0"},
	}

	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			dir := t.TempDir()
			inputPath := filepath.Join(dir, "input.pptx")

			workbook := buildWorkbookWithMissingNumeric(t, "Cat1", "Cat2", 10)
			parts := map[string][]byte{
				"ppt/slides/slide1.xml": []byte("<slide/>"),
				"ppt/slides/_rels/slide1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`),
				"ppt/charts/chart1.xml": chartWithCaches("Sheet1!$A$2:$A$3", "Sheet1!$B$2:$B$3", []string{"Old1", "Old2"}, []string{"1", "2"}),
				"ppt/charts/_rels/chart1.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/embeddedWorkbook1.xlsx"/>
</Relationships>`),
				"ppt/embeddings/embeddedWorkbook1.xlsx": workbook,
			}

			if err := writeZipFile(inputPath, parts); err != nil {
				t.Fatalf("writeZipFile: %v", err)
			}

			opts := DefaultOptions()
			opts.Workbook.MissingNumericPolicy = tc.policy
			doc, err := OpenFile(inputPath, WithOptions(opts))
			if err != nil {
				t.Fatalf("OpenFile: %v", err)
			}

			if err := doc.SyncChartCaches(); err != nil {
				t.Fatalf("SyncChartCaches: %v", err)
			}

			outputPath := filepath.Join(dir, "output.pptx")
			if err := doc.SaveFile(outputPath); err != nil {
				t.Fatalf("SaveFile: %v", err)
			}
			outChart := readZipEntry(t, outputPath, "ppt/charts/chart1.xml")
			_, nums := extractChartCacheValues(t, outChart)
			if len(nums) != 2 {
				t.Fatalf("expected 2 numeric cache values, got %d", len(nums))
			}
			if nums[1] != tc.want {
				t.Fatalf("expected missing numeric %q, got %q", tc.want, nums[1])
			}
		})
	}
}

func chartWithCaches(catFormula, valFormula string, catValues, valValues []string) []byte {
	var catBuf bytes.Buffer
	for i, v := range catValues {
		catBuf.WriteString(`<c:pt idx="`)
		catBuf.WriteString(intToString(i))
		catBuf.WriteString(`"><c:v>`)
		catBuf.WriteString(v)
		catBuf.WriteString(`</c:v></c:pt>`)
	}

	var valBuf bytes.Buffer
	for i, v := range valValues {
		valBuf.WriteString(`<c:pt idx="`)
		valBuf.WriteString(intToString(i))
		valBuf.WriteString(`"><c:v>`)
		valBuf.WriteString(v)
		valBuf.WriteString(`</c:v></c:pt>`)
	}

	xml := `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:ser>
          <c:cat><c:strRef><c:f>` + catFormula + `</c:f><c:strCache><c:ptCount val="` + intToString(len(catValues)) + `"/>` + catBuf.String() + `</c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:f>` + valFormula + `</c:f><c:numCache><c:ptCount val="` + intToString(len(valValues)) + `"/>` + valBuf.String() + `</c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`
	return []byte(xml)
}

func buildWorkbookWithValues(t *testing.T, cat1, cat2 string, val1, val2 float64) []byte {
	t.Helper()

	xml := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>` + cat1 + `</t></is></c>
      <c r="B2"><v>` + floatToString(val1) + `</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>` + cat2 + `</t></is></c>
      <c r="B3"><v>` + floatToString(val2) + `</v></c>
    </row>
  </sheetData>
</worksheet>`

	parts := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
		"xl/workbook.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
		"xl/worksheets/sheet1.xml": []byte(xml),
	}

	return writeZipBytes(t, parts)
}

func readZipEntry(t *testing.T, path, entryName string) []byte {
	t.Helper()

	file, err := os.Open(path)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	defer file.Close()

	info, err := file.Stat()
	if err != nil {
		t.Fatalf("Stat: %v", err)
	}

	reader, err := zip.NewReader(file, info.Size())
	if err != nil {
		t.Fatalf("NewReader: %v", err)
	}

	for _, part := range reader.File {
		if part.Name == entryName {
			rc, err := part.Open()
			if err != nil {
				t.Fatalf("Open entry: %v", err)
			}
			defer rc.Close()
			data, err := io.ReadAll(rc)
			if err != nil {
				t.Fatalf("Read entry: %v", err)
			}
			return data
		}
	}

	t.Fatalf("entry %q not found", entryName)
	return nil
}

func extractChartCacheValues(t *testing.T, data []byte) ([]string, []string) {
	t.Helper()

	decoder := xml.NewDecoder(bytes.NewReader(data))
	inStrCache := false
	inNumCache := false
	inValue := false
	var buf bytes.Buffer
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
					buf.Reset()
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
						cats = append(cats, buf.String())
					} else if inNumCache {
						nums = append(nums, buf.String())
					}
				}
				inValue = false
			}
		case xml.CharData:
			if inValue {
				buf.Write([]byte(tok))
			}
		}
	}

	return cats, nums
}

func intToString(value int) string {
	return strconv.FormatInt(int64(value), 10)
}

func floatToString(value float64) string {
	return strconv.FormatFloat(value, 'f', -1, 64)
}

func buildWorkbookWithMissingNumeric(t *testing.T, cat1, cat2 string, val1 float64) []byte {
	t.Helper()

	xml := `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>` + cat1 + `</t></is></c>
      <c r="B2"><v>` + floatToString(val1) + `</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>` + cat2 + `</t></is></c>
    </row>
  </sheetData>
</worksheet>`

	parts := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`),
		"xl/workbook.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
		"xl/worksheets/sheet1.xml": []byte(xml),
	}

	return writeZipBytes(t, parts)
}
