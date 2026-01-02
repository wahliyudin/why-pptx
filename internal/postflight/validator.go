package postflight

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"path"
	"sort"
	"strconv"
	"strings"

	"why-pptx/internal/overlaystage"
	"why-pptx/internal/rels"
)

type Mode string

const (
	ModeStrict     Mode = "Strict"
	ModeBestEffort Mode = "BestEffort"
)

type ValidateContext struct {
	ChartPath            string
	SlidePath            string
	WorkbookPath         string
	Mode                 Mode
	CacheSyncEnabled     bool
	MissingNumericPolicy int
}

type Document struct {
	Overlay   overlaystage.Overlay
	EmitAlert func(code, message string, ctx map[string]string)
}

type PostflightValidator struct {
	doc     *Document
	overlay overlaystage.Overlay
}

func NewPostflightValidator(doc *Document) *PostflightValidator {
	var overlay overlaystage.Overlay
	if doc != nil {
		overlay = doc.Overlay
	}
	return &PostflightValidator{
		doc:     doc,
		overlay: overlay,
	}
}

type Error struct {
	Code string
	Err  error
}

func (e *Error) Error() string {
	if e == nil {
		return "postflight error"
	}
	if e.Err != nil {
		return e.Err.Error()
	}
	return "postflight error: " + e.Code
}

func (e *Error) Unwrap() error {
	if e == nil {
		return nil
	}
	return e.Err
}

func IsPostflightError(err error) bool {
	var target *Error
	return errors.As(err, &target)
}

func (v *PostflightValidator) ValidateChartStage(ctx ValidateContext, stage *overlaystage.StagingOverlay) error {
	if stage == nil {
		return fmt.Errorf("postflight: stage is nil")
	}
	if v.overlay == nil {
		return fmt.Errorf("postflight: overlay not initialized")
	}

	touched := stage.ListTouched()
	if len(touched) == 0 {
		return nil
	}

	if err := v.checkUnexpectedParts(ctx, touched); err != nil {
		return err
	}

	touchedCharts := make([]string, 0)
	for _, part := range touched {
		if strings.HasPrefix(part, "ppt/charts/") && strings.HasSuffix(part, ".xml") {
			if err := v.checkWellFormedXML(ctx, stage, part); err != nil {
				return err
			}
			touchedCharts = append(touchedCharts, part)
		}
	}

	for _, part := range touched {
		if strings.HasPrefix(part, "ppt/embeddings/") && strings.HasSuffix(strings.ToLower(part), ".xlsx") {
			if err := v.checkSharedStrings(ctx, stage, part); err != nil {
				return err
			}
			if err := v.checkWorksheetCellTypes(ctx, stage, part); err != nil {
				return err
			}
		}
	}

	if ctx.CacheSyncEnabled {
		for _, chartPath := range touchedCharts {
			if err := v.checkChartCaches(ctx, stage, chartPath); err != nil {
				return err
			}
		}
	}

	for _, chartPath := range touchedCharts {
		if err := v.checkRelationshipTargets(ctx, stage, chartPath); err != nil {
			return err
		}
	}

	return nil
}

func (v *PostflightValidator) checkUnexpectedParts(ctx ValidateContext, touched []string) error {
	for _, part := range touched {
		exists, err := v.hasBaseline(part)
		if err != nil {
			return v.wrapError("POSTFLIGHT_UNEXPECTED_PART_ADDED", fmt.Errorf("check baseline for %q: %w", part, err), ctx, map[string]string{
				"partPath": part,
			})
		}
		if !exists {
			return v.wrapError("POSTFLIGHT_UNEXPECTED_PART_ADDED", fmt.Errorf("unexpected new part %q", part), ctx, map[string]string{
				"partPath": part,
			})
		}
	}
	return nil
}

func (v *PostflightValidator) checkWellFormedXML(ctx ValidateContext, stage *overlaystage.StagingOverlay, part string) error {
	data, err := stage.Get(part)
	if err != nil {
		return v.wrapError("POSTFLIGHT_XML_MALFORMED", fmt.Errorf("read %q: %w", part, err), ctx, map[string]string{
			"partPath": part,
		})
	}
	if err := validateXML(data); err != nil {
		return v.wrapError("POSTFLIGHT_XML_MALFORMED", fmt.Errorf("malformed xml %q: %w", part, err), ctx, map[string]string{
			"partPath": part,
		})
	}
	return nil
}

func (v *PostflightValidator) checkSharedStrings(ctx ValidateContext, stage *overlaystage.StagingOverlay, workbookPath string) error {
	data, err := stage.Get(workbookPath)
	if err != nil {
		return v.wrapError("POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED", fmt.Errorf("read workbook %q: %w", workbookPath, err), ctx, map[string]string{
			"partPath":     workbookPath,
			"workbookPath": workbookPath,
		})
	}

	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return v.wrapError("POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED", fmt.Errorf("open workbook %q: %w", workbookPath, err), ctx, map[string]string{
			"partPath":     workbookPath,
			"workbookPath": workbookPath,
		})
	}

	for _, part := range reader.File {
		if part.Name == "xl/sharedStrings.xml" {
			return v.wrapError("POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED", fmt.Errorf("workbook %q contains sharedStrings.xml", workbookPath), ctx, map[string]string{
				"partPath":     workbookPath,
				"workbookPath": workbookPath,
			})
		}
	}

	return nil
}

func (v *PostflightValidator) checkWorksheetCellTypes(ctx ValidateContext, stage *overlaystage.StagingOverlay, workbookPath string) error {
	data, err := stage.Get(workbookPath)
	if err != nil {
		return v.wrapError("POSTFLIGHT_XML_MALFORMED", fmt.Errorf("read workbook %q: %w", workbookPath, err), ctx, map[string]string{
			"partPath":     workbookPath,
			"workbookPath": workbookPath,
		})
	}

	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return v.wrapError("POSTFLIGHT_XML_MALFORMED", fmt.Errorf("open workbook %q: %w", workbookPath, err), ctx, map[string]string{
			"partPath":     workbookPath,
			"workbookPath": workbookPath,
		})
	}

	for _, part := range reader.File {
		if !strings.HasPrefix(part.Name, "xl/worksheets/") || !strings.HasSuffix(part.Name, ".xml") {
			continue
		}
		rc, err := part.Open()
		if err != nil {
			return v.wrapError("POSTFLIGHT_XML_MALFORMED", fmt.Errorf("read worksheet %q: %w", part.Name, err), ctx, map[string]string{
				"partPath":     part.Name,
				"workbookPath": workbookPath,
				"sheetPath":    part.Name,
			})
		}
		if err := v.scanWorksheetForSharedStrings(ctx, workbookPath, part.Name, rc); err != nil {
			_ = rc.Close()
			return err
		}
		_ = rc.Close()
	}

	return nil
}

func (v *PostflightValidator) scanWorksheetForSharedStrings(ctx ValidateContext, workbookPath, sheetPath string, r io.Reader) error {
	decoder := xml.NewDecoder(r)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return v.wrapError("POSTFLIGHT_XML_MALFORMED", fmt.Errorf("parse worksheet %q: %w", sheetPath, err), ctx, map[string]string{
				"partPath":     sheetPath,
				"workbookPath": workbookPath,
				"sheetPath":    sheetPath,
			})
		}

		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "c" {
			continue
		}

		var cellType string
		var cellRef string
		for _, attr := range start.Attr {
			switch attr.Name.Local {
			case "t":
				cellType = attr.Value
			case "r":
				cellRef = attr.Value
			}
		}
		if cellType == "s" {
			return v.wrapError("POSTFLIGHT_XLSX_CELL_TYPE_MISMATCH", fmt.Errorf("worksheet %q contains shared string cell", sheetPath), ctx, map[string]string{
				"partPath":     sheetPath,
				"workbookPath": workbookPath,
				"sheetPath":    sheetPath,
				"cellRef":      cellRef,
			})
		}
	}
}

func (v *PostflightValidator) checkRelationshipTargets(ctx ValidateContext, stage *overlaystage.StagingOverlay, chartPath string) error {
	relPath := chartRelsPath(chartPath)
	hasRel, err := stage.Has(relPath)
	if err != nil {
		return v.wrapError("POSTFLIGHT_REL_TARGET_MISSING", fmt.Errorf("check rels %q: %w", relPath, err), ctx, map[string]string{
			"partPath": relPath,
		})
	}
	if !hasRel {
		return nil
	}

	data, err := stage.Get(relPath)
	if err != nil {
		return v.wrapError("POSTFLIGHT_REL_TARGET_MISSING", fmt.Errorf("read rels %q: %w", relPath, err), ctx, map[string]string{
			"partPath": relPath,
		})
	}

	parsed, err := rels.Parse(bytes.NewReader(data))
	if err != nil {
		return v.wrapError("POSTFLIGHT_REL_TARGET_MISSING", fmt.Errorf("parse rels %q: %w", relPath, err), ctx, map[string]string{
			"partPath": relPath,
		})
	}

	for _, rel := range parsed.ByID {
		if rel.TargetMode == "External" {
			continue
		}
		target := resolveRelTarget(chartPath, rel.Target)
		if target == "" {
			continue
		}
		// stage.Has uses the merged view (stage overrides + parent overlay + baseline).
		exists, err := stage.Has(target)
		if err != nil {
			return v.wrapError("POSTFLIGHT_REL_TARGET_MISSING", fmt.Errorf("check rel target %q: %w", target, err), ctx, map[string]string{
				"partPath": relPath,
				"target":   target,
			})
		}
		if !exists {
			return v.wrapError("POSTFLIGHT_REL_TARGET_MISSING", fmt.Errorf("missing rel target %q", target), ctx, map[string]string{
				"partPath": relPath,
				"target":   target,
			})
		}
	}

	return nil
}

const (
	missingNumericEmpty = 0
	missingNumericZero  = 1
)

type cacheState struct {
	kind          string
	role          string
	seriesIndex   int
	plotType      string
	ptCount       int
	ptCountSeen   bool
	ptCountValue  string
	ptIdx         map[int]struct{}
	ptTotal       int
	inPt          bool
	ptHasValue    bool
	ptValue       string
	inValue       bool
	valueBuf      strings.Builder
	hasValueError bool
	inArea        bool
	values        []string
}

func (v *PostflightValidator) checkChartCaches(ctx ValidateContext, stage *overlaystage.StagingOverlay, chartPath string) error {
	data, err := stage.Get(chartPath)
	if err != nil {
		return v.wrapError("POSTFLIGHT_CHART_CACHE_INVALID", fmt.Errorf("read chart %q: %w", chartPath, err), ctx, map[string]string{
			"partPath": chartPath,
		})
	}

	decoder := xml.NewDecoder(bytes.NewReader(data))
	seriesCounter := -1
	currentSeries := -1
	serDepth := 0
	catDepth := 0
	valDepth := 0
	txDepth := 0
	barDepth := 0
	lineDepth := 0
	currentPlotType := ""
	areaDepth := 0
	areaSeries := make(map[int]struct{})
	areaCategories := make(map[int][]string)
	areaValueCounts := make(map[int]int)
	hasBarSeries := false
	hasLineSeries := false
	mixedSeries := make(map[int]struct{})
	mixedCategories := make(map[int][]string)
	mixedValueCounts := make(map[int]int)
	var cache *cacheState

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return v.wrapError("POSTFLIGHT_CHART_CACHE_INVALID", fmt.Errorf("parse chart %q: %w", chartPath, err), ctx, map[string]string{
				"partPath": chartPath,
			})
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "barChart":
				barDepth++
			case "lineChart":
				lineDepth++
			case "areaChart":
				areaDepth++
			case "ser":
				if serDepth == 0 {
					seriesCounter++
					currentSeries = seriesCounter
					if barDepth > 0 {
						currentPlotType = "bar"
						hasBarSeries = true
						mixedSeries[currentSeries] = struct{}{}
					} else if lineDepth > 0 {
						currentPlotType = "line"
						hasLineSeries = true
						mixedSeries[currentSeries] = struct{}{}
					} else {
						currentPlotType = ""
					}
					if areaDepth > 0 {
						areaSeries[currentSeries] = struct{}{}
					}
				}
				serDepth++
			case "cat":
				if serDepth > 0 {
					catDepth++
				}
			case "val":
				if serDepth > 0 {
					valDepth++
				}
			case "tx":
				if serDepth > 0 {
					txDepth++
				}
			case "strCache", "numCache":
				role := ""
				if tok.Name.Local == "strCache" {
					if catDepth > 0 {
						role = "categories"
					} else if txDepth > 0 {
						role = "seriesName"
					}
				} else if tok.Name.Local == "numCache" {
					if valDepth > 0 {
						role = "values"
					}
				}
				cache = &cacheState{
					kind:        tok.Name.Local,
					role:        role,
					seriesIndex: currentSeries,
					plotType:    currentPlotType,
					ptIdx:       make(map[int]struct{}),
					inArea:      areaDepth > 0,
				}
			case "ptCount":
				if cache != nil {
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" {
							cache.ptCountValue = attr.Value
							cache.ptCountSeen = true
							if cache.ptCount, err = parseInt(attr.Value); err != nil {
								return v.cacheError(ctx, chartPath, cache, fmt.Errorf("invalid ptCount %q", attr.Value))
							}
							break
						}
					}
					if !cache.ptCountSeen {
						return v.cacheError(ctx, chartPath, cache, fmt.Errorf("missing ptCount"))
					}
				}
			case "pt":
				if cache != nil {
					idx, err := readPtIndex(tok.Attr)
					if err != nil {
						return v.cacheError(ctx, chartPath, cache, err)
					}
					cache.ptIdx[idx] = struct{}{}
					cache.ptTotal++
					cache.inPt = true
					cache.ptHasValue = false
					cache.ptValue = ""
				}
			case "v":
				if cache != nil && cache.inPt {
					cache.inValue = true
					cache.valueBuf.Reset()
				}
			}
		case xml.EndElement:
			switch tok.Name.Local {
			case "barChart":
				if barDepth > 0 {
					barDepth--
				}
			case "lineChart":
				if lineDepth > 0 {
					lineDepth--
				}
			case "areaChart":
				if areaDepth > 0 {
					areaDepth--
				}
			case "ser":
				if serDepth > 0 {
					serDepth--
				}
				if serDepth == 0 {
					currentSeries = -1
					currentPlotType = ""
					catDepth = 0
					valDepth = 0
					txDepth = 0
				}
			case "cat":
				if catDepth > 0 {
					catDepth--
				}
			case "val":
				if valDepth > 0 {
					valDepth--
				}
			case "tx":
				if txDepth > 0 {
					txDepth--
				}
			case "v":
				if cache != nil && cache.inValue {
					cache.ptValue = cache.valueBuf.String()
					cache.ptHasValue = true
					cache.inValue = false
				}
			case "pt":
				if cache != nil && cache.inPt {
					if cache.kind == "strCache" {
						if !cache.ptHasValue {
							return v.cacheError(ctx, chartPath, cache, fmt.Errorf("missing strCache value"))
						}
						if cache.role == "categories" && (cache.inArea || cache.plotType != "") {
							cache.values = append(cache.values, cache.ptValue)
						}
					} else if cache.kind == "numCache" {
						if err := validateNumericValue(cache.ptValue, ctx.MissingNumericPolicy); err != nil {
							return v.cacheError(ctx, chartPath, cache, err)
						}
					}
					cache.inPt = false
					cache.ptHasValue = false
					cache.ptValue = ""
				}
			case "strCache", "numCache":
				if cache != nil {
					if !cache.ptCountSeen {
						return v.cacheError(ctx, chartPath, cache, fmt.Errorf("missing ptCount"))
					}
					if cache.ptTotal != cache.ptCount {
						return v.cacheError(ctx, chartPath, cache, fmt.Errorf("ptCount mismatch: expected %d got %d", cache.ptCount, cache.ptTotal))
					}
					if err := validateIdxSequence(cache.ptIdx, cache.ptCount); err != nil {
						return v.cacheError(ctx, chartPath, cache, err)
					}
					if cache.inArea {
						if cache.role == "categories" {
							areaCategories[cache.seriesIndex] = append([]string(nil), cache.values...)
						} else if cache.role == "values" {
							areaValueCounts[cache.seriesIndex] = cache.ptCount
						}
					}
					if cache.plotType != "" {
						if cache.role == "categories" {
							mixedCategories[cache.seriesIndex] = append([]string(nil), cache.values...)
						} else if cache.role == "values" {
							mixedValueCounts[cache.seriesIndex] = cache.ptCount
						}
					}
					cache = nil
				}
			}
		case xml.CharData:
			if cache != nil && cache.inValue {
				cache.valueBuf.Write([]byte(tok))
			}
		}
	}

	if len(areaSeries) > 0 {
		seriesKeys := sortedSeries(areaSeries)
		if len(areaCategories) != len(areaSeries) {
			return v.areaCacheError(ctx, chartPath, -1, fmt.Errorf("missing categories cache for area chart"))
		}
		if len(areaValueCounts) != len(areaSeries) {
			return v.areaCacheError(ctx, chartPath, -1, fmt.Errorf("missing values cache for area chart"))
		}

		baseCats := areaCategories[seriesKeys[0]]
		baseCount := areaValueCounts[seriesKeys[0]]
		for _, idx := range seriesKeys {
			cats, ok := areaCategories[idx]
			if !ok {
				return v.areaCacheError(ctx, chartPath, idx, fmt.Errorf("missing categories cache for series %d", idx))
			}
			if !equalStrings(cats, baseCats) {
				return v.areaCacheError(ctx, chartPath, idx, fmt.Errorf("area chart categories must match across series"))
			}
			if areaValueCounts[idx] != baseCount {
				return v.areaCacheError(ctx, chartPath, idx, fmt.Errorf("area chart ptCount mismatch across series"))
			}
			if len(cats) != baseCount {
				return v.areaCacheError(ctx, chartPath, idx, fmt.Errorf("area chart categories/values length mismatch"))
			}
		}
	}

	if hasBarSeries && hasLineSeries {
		if len(mixedCategories) != len(mixedSeries) {
			return v.mixedCacheError(ctx, chartPath, -1, fmt.Errorf("missing categories cache for mixed chart"))
		}
		if len(mixedValueCounts) != len(mixedSeries) {
			return v.mixedCacheError(ctx, chartPath, -1, fmt.Errorf("missing values cache for mixed chart"))
		}

		seriesKeys := sortedSeries(mixedSeries)
		baseCats := mixedCategories[seriesKeys[0]]
		baseCount := mixedValueCounts[seriesKeys[0]]
		for _, idx := range seriesKeys {
			cats, ok := mixedCategories[idx]
			if !ok {
				return v.mixedCacheError(ctx, chartPath, idx, fmt.Errorf("missing categories cache for series %d", idx))
			}
			if !equalStrings(cats, baseCats) {
				return v.mixedCacheError(ctx, chartPath, idx, fmt.Errorf("mixed chart categories must match across series"))
			}
			if mixedValueCounts[idx] != baseCount {
				return v.mixedCacheError(ctx, chartPath, idx, fmt.Errorf("mixed chart ptCount mismatch across series"))
			}
			if len(cats) != baseCount {
				return v.mixedCacheError(ctx, chartPath, idx, fmt.Errorf("mixed chart categories/values length mismatch"))
			}
		}
	}

	return nil
}

func (v *PostflightValidator) cacheError(ctx ValidateContext, chartPath string, cache *cacheState, err error) error {
	extra := map[string]string{
		"partPath": chartPath,
	}
	if cache != nil && cache.seriesIndex >= 0 {
		extra["seriesIndex"] = fmt.Sprintf("%d", cache.seriesIndex)
	}
	return v.wrapError("POSTFLIGHT_CHART_CACHE_INVALID", err, ctx, extra)
}

func (v *PostflightValidator) areaCacheError(ctx ValidateContext, chartPath string, seriesIndex int, err error) error {
	extra := map[string]string{
		"partPath": chartPath,
	}
	if seriesIndex >= 0 {
		extra["seriesIndex"] = fmt.Sprintf("%d", seriesIndex)
	}
	return v.wrapError("POSTFLIGHT_CHART_CACHE_INVALID", err, ctx, extra)
}

func (v *PostflightValidator) mixedCacheError(ctx ValidateContext, chartPath string, seriesIndex int, err error) error {
	extra := map[string]string{
		"partPath": chartPath,
	}
	if seriesIndex >= 0 {
		extra["seriesIndex"] = fmt.Sprintf("%d", seriesIndex)
	}
	return v.wrapError("POSTFLIGHT_CHART_CACHE_INVALID", err, ctx, extra)
}

func readPtIndex(attrs []xml.Attr) (int, error) {
	for _, attr := range attrs {
		if attr.Name.Local == "idx" {
			value, err := parseInt(attr.Value)
			if err != nil {
				return 0, fmt.Errorf("invalid pt idx %q", attr.Value)
			}
			if value < 0 {
				return 0, fmt.Errorf("invalid pt idx %q", attr.Value)
			}
			return value, nil
		}
	}
	return 0, fmt.Errorf("missing pt idx")
}

func parseInt(value string) (int, error) {
	out, err := strconv.Atoi(strings.TrimSpace(value))
	if err != nil {
		return 0, err
	}
	return out, nil
}

func validateIdxSequence(indices map[int]struct{}, count int) error {
	if len(indices) != count {
		return fmt.Errorf("pt idx count mismatch")
	}
	for i := 0; i < count; i++ {
		if _, ok := indices[i]; !ok {
			return fmt.Errorf("pt idx not contiguous")
		}
	}
	return nil
}

func validateNumericValue(value string, policy int) error {
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return nil
	}
	if _, err := strconv.ParseFloat(trimmed, 64); err == nil {
		return nil
	}
	if policy == missingNumericZero {
		return nil
	}
	return fmt.Errorf("invalid numeric cache value %q", trimmed)
}

func validateXML(data []byte) error {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		_, err := decoder.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
	}
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

func chartRelsPath(chartPath string) string {
	return path.Join(path.Dir(chartPath), "_rels", path.Base(chartPath)+".rels")
}

func sortedSeries(series map[int]struct{}) []int {
	keys := make([]int, 0, len(series))
	for key := range series {
		keys = append(keys, key)
	}
	sort.Ints(keys)
	return keys
}

func equalStrings(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func (v *PostflightValidator) hasBaseline(path string) (bool, error) {
	if checker, ok := v.overlay.(overlaystage.BaselineChecker); ok {
		return checker.HasBaseline(path)
	}
	return v.overlay.Has(path)
}

func (v *PostflightValidator) wrapError(code string, err error, ctx ValidateContext, extra map[string]string) error {
	v.emitAlert(code, messageForCode(code), ctx, extra)
	return &Error{Code: code, Err: err}
}

func (v *PostflightValidator) emitAlert(code, message string, ctx ValidateContext, extra map[string]string) {
	if v.doc == nil {
		return
	}
	out := make(map[string]string, 6+len(extra))
	if ctx.ChartPath != "" {
		out["chartPath"] = ctx.ChartPath
	}
	if ctx.SlidePath != "" {
		out["slidePath"] = ctx.SlidePath
	}
	if ctx.WorkbookPath != "" {
		out["workbookPath"] = ctx.WorkbookPath
	}
	if ctx.Mode != "" {
		out["mode"] = string(ctx.Mode)
	}
	out["stage"] = "postflight"
	for key, value := range extra {
		out[key] = value
	}
	if v.doc.EmitAlert == nil {
		return
	}
	v.doc.EmitAlert(code, message, out)
}

func messageForCode(code string) string {
	switch code {
	case "POSTFLIGHT_UNEXPECTED_PART_ADDED":
		return "Unexpected part added during chart update"
	case "POSTFLIGHT_XML_MALFORMED":
		return "Malformed XML detected after chart update"
	case "POSTFLIGHT_XLSX_SHAREDSTRINGS_DETECTED":
		return "Embedded workbook contains sharedStrings.xml"
	case "POSTFLIGHT_REL_TARGET_MISSING":
		return "Relationship target missing after chart update"
	case "POSTFLIGHT_CHART_CACHE_INVALID":
		return "Chart cache validation failed"
	case "POSTFLIGHT_XLSX_CELL_TYPE_MISMATCH":
		return "Worksheet uses unsupported shared string cell type"
	default:
		return "Postflight validation failed"
	}
}
