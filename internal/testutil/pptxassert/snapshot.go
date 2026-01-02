package pptxassert

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path"
	"sort"
	"strings"
	"testing"

	"why-pptx/internal/chartdiscover"
	"why-pptx/internal/chartxml"
	"why-pptx/internal/xlref"
)

type Snapshot struct {
	Entries   []string           `json:"entries"`
	Charts    []ChartSnapshot    `json:"charts"`
	Workbooks []WorkbookSnapshot `json:"workbooks"`
}

type ChartSnapshot struct {
	ChartPath    string                `json:"chartPath"`
	ChartType    string                `json:"chartType"`
	WorkbookPath string                `json:"workbookPath,omitempty"`
	Plots        []PlotSnapshot        `json:"plots,omitempty"`
	AxisGroups   []AxisGroupSnapshot   `json:"axisGroups,omitempty"`
	Series       []ChartSeriesSnapshot `json:"series"`
	Cache        CacheSnapshot         `json:"cache"`
}

type PlotSnapshot struct {
	PlotType    string   `json:"plotType"`
	AxisIDs     []string `json:"axisIds,omitempty"`
	AxisRole    string   `json:"axisRole,omitempty"`
	SeriesCount int      `json:"seriesCount"`
}

type AxisGroupSnapshot struct {
	CatAxID              string `json:"catAxId"`
	ValAxID              string `json:"valAxId"`
	ValAxisPos           string `json:"valAxisPos,omitempty"`
	ValHasMajorGridlines bool   `json:"valHasMajorGridlines,omitempty"`
	ValHasMinorGridlines bool   `json:"valHasMinorGridlines,omitempty"`
}

type ChartSeriesSnapshot struct {
	Index      int    `json:"index"`
	PlotType   string `json:"plotType,omitempty"`
	PlotIndex  int    `json:"plotIndex,omitempty"`
	Axis       string `json:"axis,omitempty"`
	Categories string `json:"categories,omitempty"`
	Values     string `json:"values,omitempty"`
	Name       string `json:"name,omitempty"`
}

type WorkbookSnapshot struct {
	WorkbookPath      string          `json:"workbookPath"`
	SharedStringsPart bool            `json:"sharedStringsPart"`
	SharedStringCells bool            `json:"sharedStringCells"`
	Sheets            []SheetSnapshot `json:"sheets"`
}

type SheetSnapshot struct {
	Sheet string         `json:"sheet"`
	Cells []CellSnapshot `json:"cells"`
}

type CellSnapshot struct {
	Ref   string `json:"ref"`
	Value string `json:"value"`
}

func BuildSnapshot(pptxPath string) (Snapshot, error) {
	entries, err := ListEntries(pptxPath)
	if err != nil {
		return Snapshot{}, err
	}

	reader, err := openZipFile(pptxPath)
	if err != nil {
		return Snapshot{}, err
	}

	pr := zipPartReader{reader: reader}
	embedded, _, err := chartdiscover.DiscoverEmbeddedCharts(pr)
	if err != nil {
		return Snapshot{}, err
	}
	workbookByChart := make(map[string]string, len(embedded))
	for _, chart := range embedded {
		workbookByChart[chart.ChartPath] = chart.WorkbookPath
	}

	chartNames := make([]string, 0)
	for _, entry := range entries {
		match, _ := path.Match("ppt/charts/*.xml", entry)
		if match && !strings.Contains(entry, "/_rels/") {
			chartNames = append(chartNames, entry)
		}
	}
	sort.Strings(chartNames)

	charts := make([]ChartSnapshot, 0, len(chartNames))
	for _, chartPath := range chartNames {
		chartXML, err := readEntry(reader, chartPath)
		if err != nil {
			return Snapshot{}, err
		}

		parsed, err := chartxml.Parse(bytes.NewReader(chartXML))
		if err != nil {
			return Snapshot{}, err
		}

		snap := ChartSnapshot{
			ChartPath:    chartPath,
			ChartType:    parsed.ChartType,
			WorkbookPath: workbookByChart[chartPath],
		}

		if parsed.ChartType == "mixed" {
			mixed, err := chartxml.ParseMixed(bytes.NewReader(chartXML))
			if err != nil {
				return Snapshot{}, err
			}

			snap.AxisGroups = axisGroupSnapshots(mixed.AxisGroups)
			snap.Plots = plotSnapshots(mixed)
			snap.Series = mixedSeriesSnapshots(mixed.Series)
		} else {
			snap.Series = seriesSnapshotsFromFormulas(parsed.Formulas)
		}

		cacheSnap, err := ExtractChartCacheSnapshot(chartXML)
		if err != nil {
			return Snapshot{}, err
		}
		snap.Cache = cacheSnap

		charts = append(charts, snap)
	}

	workbooks, err := buildWorkbookSnapshots(reader, charts)
	if err != nil {
		return Snapshot{}, err
	}

	return Snapshot{
		Entries:   entries,
		Charts:    charts,
		Workbooks: workbooks,
	}, nil
}

func WriteSnapshot(path string, snap Snapshot) error {
	data, err := json.MarshalIndent(snap, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0o644)
}

func LoadSnapshot(path string) (Snapshot, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return Snapshot{}, err
	}
	var snap Snapshot
	if err := json.Unmarshal(data, &snap); err != nil {
		return Snapshot{}, err
	}
	return snap, nil
}

func AssertSnapshotEqual(t *testing.T, got, want Snapshot) {
	t.Helper()

	if snapshotsEqual(got, want) {
		return
	}

	gotJSON, _ := json.MarshalIndent(got, "", "  ")
	wantJSON, _ := json.MarshalIndent(want, "", "  ")
	t.Fatalf("snapshot mismatch\n--- got ---\n%s\n--- want ---\n%s", string(gotJSON), string(wantJSON))
}

func snapshotsEqual(a, b Snapshot) bool {
	return jsonEqual(a, b)
}

func jsonEqual(a, b any) bool {
	dataA, err := json.Marshal(a)
	if err != nil {
		return false
	}
	dataB, err := json.Marshal(b)
	if err != nil {
		return false
	}
	return bytes.Equal(dataA, dataB)
}

type zipPartReader struct {
	reader *zip.Reader
}

func (z zipPartReader) ListParts() ([]string, error) {
	names := make([]string, 0, len(z.reader.File))
	for _, part := range z.reader.File {
		names = append(names, part.Name)
	}
	sort.Strings(names)
	return names, nil
}

func (z zipPartReader) ReadPart(name string) ([]byte, error) {
	return readEntry(z.reader, name)
}

func axisGroupSnapshots(groups []chartxml.AxisGroup) []AxisGroupSnapshot {
	out := make([]AxisGroupSnapshot, 0, len(groups))
	for _, group := range groups {
		out = append(out, AxisGroupSnapshot{
			CatAxID:              group.CatAxID,
			ValAxID:              group.ValAxID,
			ValAxisPos:           group.ValAxisPos,
			ValHasMajorGridlines: group.ValHasMajorGridlines,
			ValHasMinorGridlines: group.ValHasMinorGridlines,
		})
	}
	return out
}

func plotSnapshots(mixed *chartxml.MixedChart) []PlotSnapshot {
	out := make([]PlotSnapshot, 0, len(mixed.Plots))
	for _, plot := range mixed.Plots {
		role := ""
		for _, seriesIdx := range plot.SeriesIndices {
			for _, series := range mixed.Series {
				if series.Index == seriesIdx && series.Axis != "" {
					role = series.Axis
					break
				}
			}
			if role != "" {
				break
			}
		}
		out = append(out, PlotSnapshot{
			PlotType:    plot.PlotType,
			AxisIDs:     append([]string(nil), plot.AxisIDs...),
			AxisRole:    role,
			SeriesCount: len(plot.SeriesIndices),
		})
	}
	return out
}

func mixedSeriesSnapshots(series []chartxml.MixedSeries) []ChartSeriesSnapshot {
	out := make([]ChartSeriesSnapshot, 0, len(series))
	for _, item := range series {
		snap := ChartSeriesSnapshot{
			Index:     item.Index,
			PlotType:  item.PlotType,
			PlotIndex: item.PlotIndex,
			Axis:      item.Axis,
		}
		for _, formula := range item.Formulas {
			applyFormula(&snap, formula)
		}
		out = append(out, snap)
	}
	sort.Slice(out, func(i, j int) bool {
		return out[i].Index < out[j].Index
	})
	return out
}

func seriesSnapshotsFromFormulas(formulas []chartxml.Formula) []ChartSeriesSnapshot {
	byIndex := make(map[int]*ChartSeriesSnapshot)
	for _, formula := range formulas {
		snap, ok := byIndex[formula.SeriesIndex]
		if !ok {
			snap = &ChartSeriesSnapshot{Index: formula.SeriesIndex}
			byIndex[formula.SeriesIndex] = snap
		}
		applyFormula(snap, formula)
	}

	out := make([]ChartSeriesSnapshot, 0, len(byIndex))
	for _, snap := range byIndex {
		out = append(out, *snap)
	}
	sort.Slice(out, func(i, j int) bool {
		return out[i].Index < out[j].Index
	})
	return out
}

func applyFormula(snap *ChartSeriesSnapshot, formula chartxml.Formula) {
	switch formula.Kind {
	case chartxml.KindCategories:
		snap.Categories = formula.Formula
	case chartxml.KindValues:
		snap.Values = formula.Formula
	case chartxml.KindSeriesName:
		snap.Name = formula.Formula
	}
}

func buildWorkbookSnapshots(reader *zip.Reader, charts []ChartSnapshot) ([]WorkbookSnapshot, error) {
	targets := make(map[string]map[string]map[string]struct{})
	for _, chart := range charts {
		if chart.WorkbookPath == "" {
			continue
		}
		for _, series := range chart.Series {
			for _, formula := range []string{series.Categories, series.Values, series.Name} {
				if formula == "" {
					continue
				}
				ref, err := xlref.ParseA1Range(formula)
				if err != nil {
					return nil, fmt.Errorf("parse formula %q: %w", formula, err)
				}
				cells, err := cellsFromRange(ref.StartCell, ref.EndCell)
				if err != nil {
					return nil, fmt.Errorf("expand range %s:%s: %w", ref.StartCell, ref.EndCell, err)
				}
				addTargets(targets, chart.WorkbookPath, ref.Sheet, cells)
			}
		}
	}

	workbookPaths := make([]string, 0, len(targets))
	for workbookPath := range targets {
		workbookPaths = append(workbookPaths, workbookPath)
	}
	sort.Strings(workbookPaths)

	workbooks := make([]WorkbookSnapshot, 0, len(workbookPaths))
	for _, workbookPath := range workbookPaths {
		data, err := readEntry(reader, workbookPath)
		if err != nil {
			return nil, err
		}
		sharedStringsPart, err := hasSharedStringsPart(data)
		if err != nil {
			return nil, err
		}
		sharedStringCells, err := hasSharedStringCells(data)
		if err != nil {
			return nil, err
		}

		sheets := make([]SheetSnapshot, 0, len(targets[workbookPath]))
		sheetNames := make([]string, 0, len(targets[workbookPath]))
		for sheet := range targets[workbookPath] {
			sheetNames = append(sheetNames, sheet)
		}
		sort.Strings(sheetNames)

		for _, sheet := range sheetNames {
			cells := sortedKeysFromSet(targets[workbookPath][sheet])
			values, err := ExtractWorkbookCellSnapshot(data, sheet, cells)
			if err != nil {
				return nil, err
			}
			sheets = append(sheets, SheetSnapshot{
				Sheet: sheet,
				Cells: cellSnapshots(values),
			})
		}

		workbooks = append(workbooks, WorkbookSnapshot{
			WorkbookPath:      workbookPath,
			SharedStringsPart: sharedStringsPart,
			SharedStringCells: sharedStringCells,
			Sheets:            sheets,
		})
	}

	return workbooks, nil
}

func addTargets(targets map[string]map[string]map[string]struct{}, workbookPath, sheet string, cells []string) {
	if _, ok := targets[workbookPath]; !ok {
		targets[workbookPath] = make(map[string]map[string]struct{})
	}
	if _, ok := targets[workbookPath][sheet]; !ok {
		targets[workbookPath][sheet] = make(map[string]struct{})
	}
	for _, cell := range cells {
		targets[workbookPath][sheet][cell] = struct{}{}
	}
}

func cellsFromRange(start, end string) ([]string, error) {
	startCol, startRow, _, err := xlref.SplitCellRef(start)
	if err != nil {
		return nil, err
	}
	endCol, endRow, _, err := xlref.SplitCellRef(end)
	if err != nil {
		return nil, err
	}

	startColIdx := colToIndex(startCol)
	endColIdx := colToIndex(endCol)

	if startCol == endCol {
		if startRow > endRow {
			startRow, endRow = endRow, startRow
		}
		out := make([]string, 0, endRow-startRow+1)
		for row := startRow; row <= endRow; row++ {
			out = append(out, fmt.Sprintf("%s%d", startCol, row))
		}
		return out, nil
	}

	if startRow == endRow {
		if startColIdx > endColIdx {
			startColIdx, endColIdx = endColIdx, startColIdx
		}
		out := make([]string, 0, endColIdx-startColIdx+1)
		for col := startColIdx; col <= endColIdx; col++ {
			out = append(out, fmt.Sprintf("%s%d", indexToCol(col), startRow))
		}
		return out, nil
	}

	return nil, fmt.Errorf("non-1D range")
}

func colToIndex(col string) int {
	idx := 0
	for i := 0; i < len(col); i++ {
		ch := col[i]
		if ch < 'A' || ch > 'Z' {
			continue
		}
		idx = idx*26 + int(ch-'A'+1)
	}
	return idx
}

func indexToCol(idx int) string {
	if idx <= 0 {
		return ""
	}
	var buf []byte
	for idx > 0 {
		idx--
		buf = append(buf, byte('A'+idx%26))
		idx /= 26
	}
	for i, j := 0, len(buf)-1; i < j; i, j = i+1, j-1 {
		buf[i], buf[j] = buf[j], buf[i]
	}
	return string(buf)
}

func sortedKeysFromSet(set map[string]struct{}) []string {
	keys := make([]string, 0, len(set))
	for key := range set {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

func cellSnapshots(values map[string]string) []CellSnapshot {
	keys := make([]string, 0, len(values))
	for key := range values {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	out := make([]CellSnapshot, 0, len(keys))
	for _, key := range keys {
		out = append(out, CellSnapshot{Ref: key, Value: values[key]})
	}
	return out
}

func hasSharedStringsPart(xlsxBytes []byte) (bool, error) {
	reader, err := zip.NewReader(bytes.NewReader(xlsxBytes), int64(len(xlsxBytes)))
	if err != nil {
		return false, err
	}

	for _, part := range reader.File {
		if part.Name == "xl/sharedStrings.xml" {
			return true, nil
		}
	}
	return false, nil
}

func hasSharedStringCells(xlsxBytes []byte) (bool, error) {
	reader, err := zip.NewReader(bytes.NewReader(xlsxBytes), int64(len(xlsxBytes)))
	if err != nil {
		return false, err
	}

	for _, part := range reader.File {
		if !strings.HasPrefix(part.Name, "xl/worksheets/") || !strings.HasSuffix(part.Name, ".xml") {
			continue
		}
		rc, err := part.Open()
		if err != nil {
			return false, err
		}
		err = scanSharedStringCells(rc)
		_ = rc.Close()
		if err == nil {
			continue
		}
		if errors.Is(err, errSharedStringCell) {
			return true, nil
		}
		return false, err
	}
	return false, nil
}
