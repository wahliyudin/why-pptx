package pptx

import (
	"bytes"
	"fmt"
	"strconv"
	"strings"

	"why-pptx/internal/chartdiscover"
	"why-pptx/internal/chartxml"
)

type Range = ChartRange

type ChartDataInput map[string][]string

type PlanRequest struct {
	TargetCharts []string
	Data         ChartDataInput
	CacheSync    *bool
}

type Plan struct {
	Charts []PlannedChart `json:"charts"`
	Alerts []Alert        `json:"alerts,omitempty"`
}

type PlannedChart struct {
	Index        int     `json:"index"`
	SlidePath    string  `json:"slidePath"`
	ChartPath    string  `json:"chartPath"`
	WorkbookPath string  `json:"workbookPath"`
	ChartType    string  `json:"chartType"`
	Title        string  `json:"title,omitempty"`
	AltText      string  `json:"altText,omitempty"`
	Action       string  `json:"action"`
	ReasonCode   string  `json:"reasonCode,omitempty"`
	Dependencies []Range `json:"dependencies,omitempty"`
}

func (d *Document) Plan() (Plan, error) {
	return d.PlanChanges(PlanRequest{})
}

func (d *Document) PlanChanges(req PlanRequest) (Plan, error) {
	if d == nil || d.pkg == nil {
		return Plan{}, fmt.Errorf("document not initialized")
	}

	cacheSync := d.opts.Chart.CacheSync
	if req.CacheSync != nil {
		cacheSync = *req.CacheSync
	}

	refs, err := chartdiscover.DiscoverChartRefs(d.pkg)
	if err != nil {
		return Plan{}, err
	}

	embedded, skipped, err := chartdiscover.DiscoverEmbeddedCharts(d.pkg)
	if err != nil {
		return Plan{}, err
	}

	embeddedByPath := make(map[string]chartdiscover.EmbeddedChart, len(embedded))
	for _, item := range embedded {
		embeddedByPath[item.ChartPath] = item
	}
	skippedByPath := make(map[string]chartdiscover.SkippedChart, len(skipped))
	for _, item := range skipped {
		skippedByPath[item.ChartPath] = item
	}

	infoByPath := make(map[string]ChartInfo, len(refs))
	allInfos := make([]ChartInfo, 0, len(refs))
	var alerts []Alert

	for i, ref := range refs {
		info, infoAlerts := d.planChartInfo(i, ref, embeddedByPath[ref.ChartPath])
		allInfos = append(allInfos, info)
		infoByPath[ref.ChartPath] = info
		if len(infoAlerts) > 0 {
			alerts = append(alerts, infoAlerts...)
		}
	}

	selected, targetAlerts, err := selectPlanTargets(req.TargetCharts, allInfos, d.opts.Mode)
	if err != nil {
		plan := Plan{Charts: []PlannedChart{}, Alerts: append(alerts, targetAlerts...)}
		return plan, err
	}
	alerts = append(alerts, targetAlerts...)

	plan := Plan{Charts: make([]PlannedChart, 0, len(refs))}
	var planErr error

	for i, ref := range refs {
		if selected != nil {
			if _, ok := selected[ref.ChartPath]; !ok {
				continue
			}
		}

		info := infoByPath[ref.ChartPath]
		chart := PlannedChart{
			Index:     i,
			SlidePath: ref.SlidePath,
			ChartPath: ref.ChartPath,
			ChartType: info.ChartType,
			Title:     info.Title,
			AltText:   info.AltText,
			Action:    "apply",
		}

		if skip, ok := skippedByPath[ref.ChartPath]; ok {
			action, code, ctx := planSkipReason(skip)
			chart.Action = action
			chart.ReasonCode = code
			if code != "" {
				alerts = append(alerts, Alert{
					Level:   "warn",
					Code:    code,
					Message: planMessageForCode(code),
					Context: ctx,
				})
			}
			plan.Charts = append(plan.Charts, chart)
			continue
		}

		embeddedItem, ok := embeddedByPath[ref.ChartPath]
		if !ok {
			chart.Action = "skip"
			chart.ReasonCode = "CHART_WORKBOOK_NOT_FOUND"
			alerts = append(alerts, Alert{
				Level:   "warn",
				Code:    "CHART_WORKBOOK_NOT_FOUND",
				Message: planMessageForCode("CHART_WORKBOOK_NOT_FOUND"),
				Context: map[string]string{
					"slide": ref.SlidePath,
					"chart": ref.ChartPath,
				},
			})
			plan.Charts = append(plan.Charts, chart)
			continue
		}

		chart.WorkbookPath = embeddedItem.WorkbookPath

		deps, err := d.extractChartDependencies(EmbeddedChart{
			SlidePath:    embeddedItem.SlidePath,
			ChartPath:    embeddedItem.ChartPath,
			WorkbookPath: embeddedItem.WorkbookPath,
		})
		if err != nil {
			chart.Action = "skip"
			chart.ReasonCode = "CHART_DEPENDENCIES_PARSE_FAILED"
			alerts = append(alerts, Alert{
				Level:   "warn",
				Code:    "CHART_DEPENDENCIES_PARSE_FAILED",
				Message: planMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
				Context: map[string]string{
					"slide":    embeddedItem.SlidePath,
					"chart":    embeddedItem.ChartPath,
					"workbook": embeddedItem.WorkbookPath,
					"error":    err.Error(),
				},
			})
			if d.opts.Mode == Strict && planErr == nil {
				planErr = err
			}
			plan.Charts = append(plan.Charts, chart)
			continue
		}

		chart.ChartType = deps.ChartType
		chart.Dependencies = deps.Ranges

		if err := validatePlanRanges(chart.Dependencies); err != nil {
			chart.Action = "skip"
			chart.ReasonCode = "CHART_DEPENDENCIES_PARSE_FAILED"
			alerts = append(alerts, Alert{
				Level:   "warn",
				Code:    "CHART_DEPENDENCIES_PARSE_FAILED",
				Message: planMessageForCode("CHART_DEPENDENCIES_PARSE_FAILED"),
				Context: map[string]string{
					"slide":    embeddedItem.SlidePath,
					"chart":    embeddedItem.ChartPath,
					"workbook": embeddedItem.WorkbookPath,
					"error":    err.Error(),
				},
			})
			if d.opts.Mode == Strict && planErr == nil {
				planErr = err
			}
			plan.Charts = append(plan.Charts, chart)
			continue
		}

		if deps.ChartType != "bar" && deps.ChartType != "line" && cacheSync {
			chart.Action = "unsupported"
			chart.ReasonCode = "CHART_TYPE_UNSUPPORTED"
			alerts = append(alerts, Alert{
				Level:   "warn",
				Code:    "CHART_TYPE_UNSUPPORTED",
				Message: planMessageForCode("CHART_TYPE_UNSUPPORTED"),
				Context: map[string]string{
					"slide":     embeddedItem.SlidePath,
					"chart":     embeddedItem.ChartPath,
					"chartType": deps.ChartType,
				},
			})
			if d.opts.Mode == Strict && planErr == nil {
				planErr = fmt.Errorf("unsupported chart type %q", deps.ChartType)
			}
			plan.Charts = append(plan.Charts, chart)
			continue
		}

		if len(req.Data) > 0 {
			action, reason, dataAlerts, dataErr := validatePlanData(req.Data, chart, d.opts.Mode)
			if len(dataAlerts) > 0 {
				alerts = append(alerts, dataAlerts...)
			}
			if dataErr != nil {
				if planErr == nil {
					planErr = dataErr
				}
				chart.Action = "skip"
				chart.ReasonCode = reason
				plan.Charts = append(plan.Charts, chart)
				continue
			}
			if action != "" {
				chart.Action = action
				chart.ReasonCode = reason
				plan.Charts = append(plan.Charts, chart)
				continue
			}
		}

		plan.Charts = append(plan.Charts, chart)
	}

	if len(alerts) > 0 {
		plan.Alerts = alerts
	}
	return plan, planErr
}

func (d *Document) planChartInfo(index int, ref chartdiscover.ChartRef, embedded chartdiscover.EmbeddedChart) (ChartInfo, []Alert) {
	info := ChartInfo{
		Index:        index,
		SlidePath:    ref.SlidePath,
		ChartPath:    ref.ChartPath,
		WorkbookPath: embedded.WorkbookPath,
		ChartType:    "unknown",
	}

	titleFromSlide, altText := d.slideChartAltText(ref.SlidePath, ref.ChartPath)
	info.AltText = altText

	data, err := d.pkg.ReadPart(ref.ChartPath)
	if err != nil {
		if titleFromSlide != "" {
			info.Title = titleFromSlide
		}
		return info, []Alert{{
			Level:   "warn",
			Code:    "CHART_INFO_PARSE_FAILED",
			Message: planMessageForCode("CHART_INFO_PARSE_FAILED"),
			Context: map[string]string{
				"slide": ref.SlidePath,
				"chart": ref.ChartPath,
				"error": err.Error(),
			},
		}}
	}

	parsed, err := chartxml.ParseInfo(bytes.NewReader(data))
	if err != nil {
		if titleFromSlide != "" {
			info.Title = titleFromSlide
		}
		return info, []Alert{{
			Level:   "warn",
			Code:    "CHART_INFO_PARSE_FAILED",
			Message: planMessageForCode("CHART_INFO_PARSE_FAILED"),
			Context: map[string]string{
				"slide": ref.SlidePath,
				"chart": ref.ChartPath,
				"error": err.Error(),
			},
		}}
	}

	info.ChartType = parsed.ChartType
	info.SeriesCount = parsed.SeriesCount
	info.Title = parsed.Title
	if info.Title == "" && titleFromSlide != "" {
		info.Title = titleFromSlide
	}

	return info, nil
}

func selectPlanTargets(targets []string, infos []ChartInfo, mode ErrorMode) (map[string]struct{}, []Alert, error) {
	if len(targets) == 0 {
		return nil, nil, nil
	}

	selected := make(map[string]struct{})
	var alerts []Alert

	pathIndex := make(map[string]ChartInfo, len(infos))
	for _, info := range infos {
		pathIndex[info.ChartPath] = info
	}

	for _, target := range targets {
		if info, ok := pathIndex[target]; ok {
			selected[info.ChartPath] = struct{}{}
			continue
		}

		matches := matchChartsByName(infos, target)
		if len(matches) == 0 {
			return nil, alerts, fmt.Errorf("chart not found")
		}
		if len(matches) > 1 {
			if mode == BestEffort {
				alerts = append(alerts, Alert{
					Level:   "warn",
					Code:    "CHART_NAME_AMBIGUOUS",
					Message: planMessageForCode("CHART_NAME_AMBIGUOUS"),
					Context: map[string]string{
						"name":    target,
						"matches": strconv.Itoa(len(matches)),
					},
				})
				continue
			}
			return nil, alerts, fmt.Errorf("ambiguous chart name")
		}
		selected[matches[0].ChartPath] = struct{}{}
	}

	return selected, alerts, nil
}

func planSkipReason(skip chartdiscover.SkippedChart) (string, string, map[string]string) {
	switch skip.Reason {
	case chartdiscover.ReasonLinked:
		return "linked", "CHART_LINKED_WORKBOOK", map[string]string{
			"slide":  skip.SlidePath,
			"chart":  skip.ChartPath,
			"target": skip.Target,
		}
	case chartdiscover.ReasonRelsMissing:
		return "skip", "CHART_RELS_MISSING", map[string]string{
			"slide":     skip.SlidePath,
			"chart":     skip.ChartPath,
			"rels_path": skip.RelsPath,
		}
	case chartdiscover.ReasonWorkbookNotFound:
		return "skip", "CHART_WORKBOOK_NOT_FOUND", map[string]string{
			"slide": skip.SlidePath,
			"chart": skip.ChartPath,
		}
	case chartdiscover.ReasonUnsupported:
		return "skip", "CHART_WORKBOOK_UNSUPPORTED_TARGET", map[string]string{
			"slide":  skip.SlidePath,
			"chart":  skip.ChartPath,
			"target": skip.Target,
		}
	default:
		return "skip", "", map[string]string{
			"slide": skip.SlidePath,
			"chart": skip.ChartPath,
		}
	}
}

func validatePlanRanges(ranges []Range) error {
	for _, r := range ranges {
		if _, err := expandRangeCells(r.StartCell, r.EndCell); err != nil {
			return err
		}
	}
	return nil
}

func validatePlanData(data ChartDataInput, chart PlannedChart, mode ErrorMode) (string, string, []Alert, error) {
	categories, hasCategories := data["categories"]
	if hasCategories {
		categoriesLen := len(categories)
		for _, r := range chart.Dependencies {
			if r.Kind != RangeValues {
				continue
			}
			key := fmt.Sprintf("values:%d", r.SeriesIndex)
			values, ok := data[key]
			if !ok {
				continue
			}
			if len(values) != categoriesLen {
				if mode == BestEffort {
					return "skip", "CHART_DATA_LENGTH_MISMATCH", []Alert{{
						Level:   "warn",
						Code:    "CHART_DATA_LENGTH_MISMATCH",
						Message: planMessageForCode("CHART_DATA_LENGTH_MISMATCH"),
						Context: map[string]string{
							"chartIndex":    strconv.Itoa(chart.Index),
							"categoriesLen": strconv.Itoa(categoriesLen),
							"valuesLen":     strconv.Itoa(len(values)),
							"seriesIndex":   strconv.Itoa(r.SeriesIndex),
						},
					}}, nil
				}
				return "", "", nil, fmt.Errorf("categories length %d does not match values length %d for series %d", categoriesLen, len(values), r.SeriesIndex)
			}
		}
	}

	for _, r := range chart.Dependencies {
		switch r.Kind {
		case RangeCategories:
			if !hasCategories {
				return "", "", nil, fmt.Errorf("categories data is required")
			}
			cells, err := expandRangeCells(r.StartCell, r.EndCell)
			if err != nil {
				return "", "", nil, err
			}
			if len(categories) != len(cells) {
				return "", "", nil, fmt.Errorf("categories length mismatch: expected %d got %d", len(cells), len(categories))
			}
		case RangeValues:
			key := fmt.Sprintf("values:%d", r.SeriesIndex)
			values, ok := data[key]
			if !ok {
				return "", "", nil, fmt.Errorf("values data missing for series %d", r.SeriesIndex)
			}
			cells, err := expandRangeCells(r.StartCell, r.EndCell)
			if err != nil {
				return "", "", nil, err
			}
			if len(values) != len(cells) {
				return "", "", nil, fmt.Errorf("values length mismatch for series %d: expected %d got %d", r.SeriesIndex, len(cells), len(values))
			}
			for _, value := range values {
				if _, err := strconv.ParseFloat(strings.TrimSpace(value), 64); err != nil {
					return "", "", nil, fmt.Errorf("invalid numeric value %q for series %d", value, r.SeriesIndex)
				}
			}
		}
	}

	return "", "", nil, nil
}

func planMessageForCode(code string) string {
	switch code {
	case "CHART_LINKED_WORKBOOK":
		return "Chart uses linked workbook and is skipped"
	case "CHART_RELS_MISSING":
		return "Chart relationships file is missing; chart is skipped"
	case "CHART_WORKBOOK_NOT_FOUND":
		return "No workbook relationship found for chart; chart is skipped"
	case "CHART_WORKBOOK_UNSUPPORTED_TARGET":
		return "Chart workbook target is unsupported; chart is skipped"
	case "CHART_DEPENDENCIES_PARSE_FAILED":
		return "Failed to extract chart dependencies; chart is skipped"
	case "CHART_CACHE_SYNC_FAILED":
		return "Failed to sync chart caches; chart is skipped"
	case "CHART_DATA_LENGTH_MISMATCH":
		return "Categories and values length mismatch; chart skipped"
	case "CHART_NAME_AMBIGUOUS":
		return "Chart name is ambiguous; no chart selected"
	case "CHART_INFO_PARSE_FAILED":
		return "Failed to parse chart info; chart metadata is partial"
	case "CHART_TYPE_UNSUPPORTED":
		return "Chart type is unsupported; chart is skipped"
	default:
		return "Plan detected an issue"
	}
}
