package pptxassert

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"path"
	"sort"
	"strconv"
	"strings"
	"testing"

	"why-pptx/internal/rels"
	"why-pptx/internal/xlref"
)

type CachePoint struct {
	Idx   int
	Value string
}

type CacheSeries struct {
	Kind        string
	SeriesIndex int
	PtCount     int
	Points      []CachePoint
}

type CacheSnapshot struct {
	Series []CacheSeries
}

type ExpectedCacheSeries struct {
	Kind        string
	SeriesIndex int
	Values      []string
}

type ExpectedCache struct {
	Series []ExpectedCacheSeries
}

type ManifestOpt func(*manifestOptions)

type manifestOptions struct {
	allowAdded   map[string]struct{}
	allowRemoved map[string]struct{}
}

func AllowAdded(names ...string) ManifestOpt {
	return func(opts *manifestOptions) {
		if opts.allowAdded == nil {
			opts.allowAdded = make(map[string]struct{}, len(names))
		}
		for _, name := range names {
			opts.allowAdded[name] = struct{}{}
		}
	}
}

func AllowRemoved(names ...string) ManifestOpt {
	return func(opts *manifestOptions) {
		if opts.allowRemoved == nil {
			opts.allowRemoved = make(map[string]struct{}, len(names))
		}
		for _, name := range names {
			opts.allowRemoved[name] = struct{}{}
		}
	}
}

func ListEntries(pptxPath string) ([]string, error) {
	reader, err := openZipFile(pptxPath)
	if err != nil {
		return nil, err
	}
	names := make([]string, 0, len(reader.File))
	for _, part := range reader.File {
		names = append(names, part.Name)
	}
	sort.Strings(names)
	return names, nil
}

func ReadEntry(pptxPath, entryName string) ([]byte, error) {
	reader, err := openZipFile(pptxPath)
	if err != nil {
		return nil, err
	}
	return readEntry(reader, entryName)
}

func AssertSameEntrySet(t *testing.T, before, after string, opts ...ManifestOpt) {
	t.Helper()

	options := manifestOptions{}
	for _, opt := range opts {
		if opt != nil {
			opt(&options)
		}
	}

	beforeEntries, err := ListEntries(before)
	if err != nil {
		t.Fatalf("ListEntries before: %v", err)
	}
	afterEntries, err := ListEntries(after)
	if err != nil {
		t.Fatalf("ListEntries after: %v", err)
	}

	beforeSet := toSet(beforeEntries)
	afterSet := toSet(afterEntries)

	var missing []string
	for name := range beforeSet {
		if _, ok := afterSet[name]; ok {
			continue
		}
		if _, ok := options.allowRemoved[name]; ok {
			continue
		}
		missing = append(missing, name)
	}

	var extra []string
	for name := range afterSet {
		if _, ok := beforeSet[name]; ok {
			continue
		}
		if _, ok := options.allowAdded[name]; ok {
			continue
		}
		extra = append(extra, name)
	}

	sort.Strings(missing)
	sort.Strings(extra)

	if len(missing) > 0 || len(extra) > 0 {
		t.Fatalf("entry set mismatch: missing=%v extra=%v", missing, extra)
	}
}

func ListRelTargets(pptxPath, relPath string) ([]string, error) {
	reader, err := openZipFile(pptxPath)
	if err != nil {
		return nil, err
	}
	data, err := readEntry(reader, relPath)
	if err != nil {
		return nil, err
	}

	parsed, err := rels.Parse(bytes.NewReader(data))
	if err != nil {
		return nil, err
	}

	basePart := basePartFromRelPath(relPath)
	targets := make([]string, 0, len(parsed.ByID))
	for _, rel := range parsed.ByID {
		if rel.TargetMode == "External" {
			continue
		}
		target := resolveRelTarget(basePart, rel.Target)
		if target == "" {
			continue
		}
		targets = append(targets, target)
	}
	sort.Strings(targets)
	return targets, nil
}

func AssertRelTargetsExist(t *testing.T, pptxPath string, relPaths []string) {
	t.Helper()

	reader, err := openZipFile(pptxPath)
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}

	entrySet := make(map[string]struct{}, len(reader.File))
	for _, part := range reader.File {
		entrySet[part.Name] = struct{}{}
	}

	for _, relPath := range relPaths {
		targets, err := ListRelTargets(pptxPath, relPath)
		if err != nil {
			t.Fatalf("ListRelTargets %s: %v", relPath, err)
		}
		for _, target := range targets {
			if _, ok := entrySet[target]; !ok {
				t.Fatalf("missing relationship target %q referenced by %s", target, relPath)
			}
		}
	}
}

func ExtractChartCacheSnapshot(chartXML []byte) (CacheSnapshot, error) {
	decoder := xml.NewDecoder(bytes.NewReader(chartXML))
	seriesCounter := -1
	currentSeries := -1
	serDepth := 0
	var snapshot CacheSnapshot

	var current *CacheSeries
	inPt := false
	var ptIdx int
	var inValue bool
	var valueBuf strings.Builder

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return CacheSnapshot{}, err
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "ser":
				if serDepth == 0 {
					seriesCounter++
					currentSeries = seriesCounter
				}
				serDepth++
			case "strCache", "numCache":
				current = &CacheSeries{
					Kind:        tok.Name.Local,
					SeriesIndex: currentSeries,
					PtCount:     -1,
				}
			case "ptCount":
				if current != nil {
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" {
							if val, err := parseInt(attr.Value); err == nil {
								current.PtCount = val
							}
							break
						}
					}
				}
			case "pt":
				if current != nil {
					idx, err := readIdxAttr(tok.Attr)
					if err == nil {
						ptIdx = idx
					} else {
						ptIdx = -1
					}
					inPt = true
				}
			case "v":
				if current != nil && inPt {
					inValue = true
					valueBuf.Reset()
				}
			}
		case xml.EndElement:
			switch tok.Name.Local {
			case "ser":
				if serDepth > 0 {
					serDepth--
				}
				if serDepth == 0 {
					currentSeries = -1
				}
			case "v":
				if inValue {
					inValue = false
				}
			case "pt":
				if current != nil && inPt {
					current.Points = append(current.Points, CachePoint{Idx: ptIdx, Value: valueBuf.String()})
					inPt = false
					valueBuf.Reset()
				}
			case "strCache", "numCache":
				if current != nil {
					snapshot.Series = append(snapshot.Series, *current)
					current = nil
				}
			}
		case xml.CharData:
			if inValue {
				valueBuf.Write([]byte(tok))
			}
		}
	}

	return snapshot, nil
}

func AssertCacheMatchesExpected(t *testing.T, snap CacheSnapshot, expected ExpectedCache) {
	t.Helper()

	if len(snap.Series) != len(expected.Series) {
		t.Fatalf("cache series count mismatch: got %d want %d", len(snap.Series), len(expected.Series))
	}

	seriesByKey := make(map[string]CacheSeries, len(snap.Series))
	for _, series := range snap.Series {
		key := fmt.Sprintf("%s:%d", series.Kind, series.SeriesIndex)
		seriesByKey[key] = series
	}

	for _, exp := range expected.Series {
		key := fmt.Sprintf("%s:%d", exp.Kind, exp.SeriesIndex)
		series, ok := seriesByKey[key]
		if !ok {
			t.Fatalf("cache series %s not found", key)
		}

		if series.PtCount != len(exp.Values) {
			t.Fatalf("cache series %s ptCount=%d want %d", key, series.PtCount, len(exp.Values))
		}

		valuesByIdx := make(map[int]string, len(series.Points))
		for _, pt := range series.Points {
			valuesByIdx[pt.Idx] = pt.Value
		}

		for i, want := range exp.Values {
			got, ok := valuesByIdx[i]
			if !ok {
				t.Fatalf("cache series %s missing idx %d", key, i)
			}
			if got != want {
				t.Fatalf("cache series %s idx %d value %q want %q", key, i, got, want)
			}
		}
	}
}

func ExtractWorkbookCellSnapshot(xlsxBytes []byte, sheetNameOrPath string, cells []string) (map[string]string, error) {
	reader, err := zip.NewReader(bytes.NewReader(xlsxBytes), int64(len(xlsxBytes)))
	if err != nil {
		return nil, err
	}
	index := make(map[string]*zip.File, len(reader.File))
	for _, part := range reader.File {
		index[part.Name] = part
	}

	sheetPath := sheetNameOrPath
	if !strings.Contains(sheetNameOrPath, "/") && !strings.HasSuffix(sheetNameOrPath, ".xml") {
		resolved, err := resolveSheetPath(index, sheetNameOrPath)
		if err != nil {
			return nil, err
		}
		sheetPath = resolved
	}

	part, ok := index[sheetPath]
	if !ok {
		return nil, fmt.Errorf("sheet %q not found", sheetPath)
	}

	targets := make(map[string]struct{}, len(cells))
	ordered := make([]string, 0, len(cells))
	for _, cell := range cells {
		normalized, err := xlref.NormalizeCellRef(cell)
		if err != nil {
			return nil, err
		}
		if _, ok := targets[normalized]; ok {
			continue
		}
		targets[normalized] = struct{}{}
		ordered = append(ordered, normalized)
	}

	rc, err := part.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	values, err := readCellValues(rc, targets)
	if err != nil {
		return nil, err
	}

	out := make(map[string]string, len(ordered))
	for _, ref := range ordered {
		out[ref] = values[ref]
	}
	return out, nil
}

func AssertNoSharedStringsPart(t *testing.T, xlsxBytes []byte) {
	t.Helper()

	reader, err := zip.NewReader(bytes.NewReader(xlsxBytes), int64(len(xlsxBytes)))
	if err != nil {
		t.Fatalf("open xlsx: %v", err)
	}

	for _, part := range reader.File {
		if part.Name == "xl/sharedStrings.xml" {
			t.Fatalf("sharedStrings.xml present in workbook")
		}
	}
}

func AssertNoSharedStringCells(t *testing.T, xlsxBytes []byte) {
	t.Helper()

	reader, err := zip.NewReader(bytes.NewReader(xlsxBytes), int64(len(xlsxBytes)))
	if err != nil {
		t.Fatalf("open xlsx: %v", err)
	}

	for _, part := range reader.File {
		if !strings.HasPrefix(part.Name, "xl/worksheets/") || !strings.HasSuffix(part.Name, ".xml") {
			continue
		}
		rc, err := part.Open()
		if err != nil {
			t.Fatalf("open worksheet %q: %v", part.Name, err)
		}
		if err := scanSharedStringCells(rc); err != nil {
			_ = rc.Close()
			t.Fatalf("%v", err)
		}
		_ = rc.Close()
	}
}

func openZipFile(filePath string) (*zip.Reader, error) {
	data, err := os.ReadFile(filePath)
	if err != nil {
		return nil, err
	}
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	return reader, nil
}

func readEntry(reader *zip.Reader, entryName string) ([]byte, error) {
	for _, part := range reader.File {
		if part.Name == entryName {
			rc, err := part.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			data, err := io.ReadAll(rc)
			if err != nil {
				return nil, err
			}
			return data, nil
		}
	}
	return nil, fmt.Errorf("entry %q not found", entryName)
}

func toSet(values []string) map[string]struct{} {
	out := make(map[string]struct{}, len(values))
	for _, value := range values {
		out[value] = struct{}{}
	}
	return out
}

func basePartFromRelPath(relPath string) string {
	dir := path.Dir(relPath)
	parent := path.Dir(dir)
	base := strings.TrimSuffix(path.Base(relPath), ".rels")
	return path.Join(parent, base)
}

func resolveRelTarget(basePart, relTarget string) string {
	if relTarget == "" {
		return ""
	}
	if strings.HasPrefix(relTarget, "/") {
		return strings.TrimLeft(relTarget, "/")
	}
	return rels.ResolveTarget(basePart, relTarget)
}

func parseInt(value string) (int, error) {
	out, err := strconv.Atoi(strings.TrimSpace(value))
	if err != nil {
		return 0, err
	}
	return out, nil
}

func readIdxAttr(attrs []xml.Attr) (int, error) {
	for _, attr := range attrs {
		if attr.Name.Local == "idx" {
			return parseInt(attr.Value)
		}
	}
	return 0, fmt.Errorf("missing idx")
}

func resolveSheetPath(index map[string]*zip.File, sheetName string) (string, error) {
	workbook, ok := index["xl/workbook.xml"]
	if !ok {
		return "", fmt.Errorf("workbook.xml not found")
	}
	relsPart, ok := index["xl/_rels/workbook.xml.rels"]
	if !ok {
		return "", fmt.Errorf("workbook rels not found")
	}

	workbookData, err := readZipPart(workbook)
	if err != nil {
		return "", err
	}
	relsData, err := readZipPart(relsPart)
	if err != nil {
		return "", err
	}

	sheets, err := parseWorkbookSheets(workbookData)
	if err != nil {
		return "", err
	}
	relID, ok := sheets[sheetName]
	if !ok {
		return "", fmt.Errorf("sheet %q not found", sheetName)
	}

	parsed, err := rels.Parse(bytes.NewReader(relsData))
	if err != nil {
		return "", err
	}
	rel, ok := parsed.Resolve(relID)
	if !ok {
		return "", fmt.Errorf("sheet %q missing rel %q", sheetName, relID)
	}

	target := rels.ResolveTarget("xl/workbook.xml", rel.Target)
	target = path.Clean(target)
	if !strings.HasPrefix(target, "xl/") {
		target = path.Join("xl", strings.TrimLeft(target, "/"))
	}
	return target, nil
}

func parseWorkbookSheets(data []byte) (map[string]string, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	sheets := make(map[string]string)

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "sheet" {
			continue
		}

		var name string
		var relID string
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "name":
				name = attr.Value
			case "id":
				if attr.Name.Space != "" {
					relID = attr.Value
				}
			}
		}
		if name != "" && relID != "" {
			sheets[name] = relID
		}
	}

	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found")
	}
	return sheets, nil
}

func readZipPart(part *zip.File) ([]byte, error) {
	rc, err := part.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

func readCellValues(r io.Reader, targets map[string]struct{}) (map[string]string, error) {
	decoder := xml.NewDecoder(r)
	values := make(map[string]string, len(targets))

	var inCell bool
	var cellRef string
	var cellType string
	var inInlineStr bool
	var inValue bool
	var hasValue bool
	var valueBuf strings.Builder

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "c":
				cellRef = ""
				cellType = ""
				inCell = false
				inInlineStr = false
				hasValue = false
				for _, attr := range tok.Attr {
					if attr.Name.Local == "r" {
						cellRef = attr.Value
					} else if attr.Name.Local == "t" {
						cellType = attr.Value
					}
				}
				if cellRef != "" {
					normalized, err := xlref.NormalizeCellRef(cellRef)
					if err == nil {
						if _, ok := targets[normalized]; ok {
							cellRef = normalized
							inCell = true
							valueBuf.Reset()
							if cellType == "inlineStr" {
								inInlineStr = true
							}
						}
					}
				}
			case "v":
				if inCell && (cellType == "" || cellType == "n") {
					inValue = true
					valueBuf.Reset()
					hasValue = true
				}
			case "t":
				if inInlineStr {
					inValue = true
					hasValue = true
				}
			}
		case xml.EndElement:
			switch tok.Name.Local {
			case "c":
				if inCell {
					if cellType == "inlineStr" {
						values[cellRef] = valueBuf.String()
					} else if hasValue {
						values[cellRef] = strings.TrimSpace(valueBuf.String())
					}
				}
				inCell = false
				inInlineStr = false
				inValue = false
				hasValue = false
				cellRef = ""
				cellType = ""
			case "v":
				inValue = false
			case "t":
				inValue = false
			}
		case xml.CharData:
			if inCell && inValue {
				valueBuf.Write([]byte(tok))
			}
		}
	}

	for ref := range targets {
		if _, ok := values[ref]; !ok {
			values[ref] = ""
		}
	}

	return values, nil
}

var errSharedStringCell = errors.New("shared string cell detected")

func scanSharedStringCells(r io.Reader) error {
	decoder := xml.NewDecoder(r)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}

		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "c" {
			continue
		}

		for _, attr := range start.Attr {
			if attr.Name.Local == "t" && attr.Value == "s" {
				return errSharedStringCell
			}
		}
	}
}
