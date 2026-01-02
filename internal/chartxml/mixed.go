package chartxml

import (
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strings"
)

type MixedSeries struct {
	Index     int
	PlotType  string
	PlotIndex int
	Axis      string
	Formulas  []Formula
}

type AxisGroup struct {
	CatAxID              string
	ValAxID              string
	ValAxisPos           string
	ValHasMajorGridlines bool
	ValHasMinorGridlines bool
}

type MixedPlot struct {
	PlotType      string
	AxisIDs       []string
	SeriesIndices []int
}

type MixedChart struct {
	Series     []MixedSeries
	Plots      []MixedPlot
	AxisGroups []AxisGroup
}

func ParseMixed(r io.Reader) (*MixedChart, error) {
	decoder := xml.NewDecoder(r)
	out := &MixedChart{}

	plotTypes := make(map[string]struct{})
	var plots []*plotState
	currentPlot := -1
	barDepth := 0
	lineDepth := 0

	axisDepth := 0
	axisKind := ""
	currentAxis := axisInfo{}
	axes := make([]axisInfo, 0)

	seriesIndex := -1
	serDepth := 0
	currentSeries := -1
	catDepth := 0
	valDepth := 0
	txDepth := 0

	inFormula := false
	formulaKind := ""
	var buf strings.Builder

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse mixed chart: %w", err)
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "barChart":
				barDepth++
				if barDepth == 1 {
					plots = append(plots, newPlotState("bar"))
					currentPlot = len(plots) - 1
					plotTypes["bar"] = struct{}{}
				}
			case "lineChart":
				lineDepth++
				if lineDepth == 1 {
					plots = append(plots, newPlotState("line"))
					currentPlot = len(plots) - 1
					plotTypes["line"] = struct{}{}
				}
			case "catAx":
				axisDepth++
				if axisDepth == 1 {
					axisKind = "cat"
					currentAxis = axisInfo{kind: "cat"}
				}
			case "valAx":
				axisDepth++
				if axisDepth == 1 {
					axisKind = "val"
					currentAxis = axisInfo{kind: "val"}
				}
			default:
				if isUnsupportedPlot(tok.Name.Local) {
					return nil, fmt.Errorf("unsupported plot type %q", tok.Name.Local)
				}
			}

			if axisDepth > 0 {
				switch tok.Name.Local {
				case "axId":
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" && currentAxis.id == "" {
							currentAxis.id = attr.Value
						}
					}
				case "crossAx":
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" {
							currentAxis.cross = attr.Value
						}
					}
				case "axisPos":
					if axisKind == "val" {
						for _, attr := range tok.Attr {
							if attr.Name.Local == "val" {
								currentAxis.axisPos = attr.Value
							}
						}
					}
				case "majorGridlines":
					if axisKind == "val" {
						currentAxis.hasMajorGridlines = true
					}
				case "minorGridlines":
					if axisKind == "val" {
						currentAxis.hasMinorGridlines = true
					}
				}
			}

			if currentPlot >= 0 {
				switch tok.Name.Local {
				case "grouping":
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" {
							if attr.Value == "stacked" || attr.Value == "percentStacked" {
								return nil, fmt.Errorf("stacked charts are unsupported")
							}
						}
					}
				case "axId":
					for _, attr := range tok.Attr {
						if attr.Name.Local == "val" && attr.Value != "" {
							plots[currentPlot].axisIDs[attr.Value] = struct{}{}
						}
					}
				case "ser":
					if serDepth == 0 {
						seriesIndex++
						plotSeriesIndex := plots[currentPlot].seriesCount
						plots[currentPlot].seriesCount++
						out.Series = append(out.Series, MixedSeries{
							Index:     seriesIndex,
							PlotType:  plots[currentPlot].plotType,
							PlotIndex: plotSeriesIndex,
						})
						currentSeries = len(out.Series) - 1
						plots[currentPlot].seriesIndices = append(plots[currentPlot].seriesIndices, seriesIndex)
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
				case "f":
					if serDepth > 0 {
						kind := ""
						if catDepth > 0 {
							kind = KindCategories
						} else if valDepth > 0 {
							kind = KindValues
						} else if txDepth > 0 {
							kind = KindSeriesName
						}
						if kind != "" && currentSeries >= 0 {
							inFormula = true
							formulaKind = kind
							buf.Reset()
						}
					}
				}
			}
		case xml.EndElement:
			switch tok.Name.Local {
			case "barChart":
				if barDepth > 0 {
					barDepth--
				}
				if barDepth == 0 && lineDepth == 0 {
					currentPlot = -1
				}
			case "lineChart":
				if lineDepth > 0 {
					lineDepth--
				}
				if barDepth == 0 && lineDepth == 0 {
					currentPlot = -1
				}
			case "catAx", "valAx":
				if axisDepth > 0 {
					axisDepth--
				}
				if axisDepth == 0 {
					if currentAxis.id != "" {
						axes = append(axes, currentAxis)
					}
					axisKind = ""
					currentAxis = axisInfo{}
				}
			case "ser":
				if serDepth > 0 {
					serDepth--
				}
				if serDepth == 0 {
					currentSeries = -1
					catDepth = 0
					valDepth = 0
					txDepth = 0
					inFormula = false
					formulaKind = ""
					buf.Reset()
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
			case "f":
				if inFormula && currentSeries >= 0 {
					text := strings.TrimSpace(buf.String())
					if text != "" {
						out.Series[currentSeries].Formulas = append(out.Series[currentSeries].Formulas, Formula{
							Kind:        formulaKind,
							SeriesIndex: out.Series[currentSeries].Index,
							Formula:     text,
						})
					}
					inFormula = false
					formulaKind = ""
					buf.Reset()
				}
			}
		case xml.CharData:
			if inFormula {
				buf.Write([]byte(tok))
			}
		}
	}

	if len(plotTypes) == 0 {
		return nil, fmt.Errorf("no supported plot types found")
	}
	if len(plotTypes) != 2 || !hasPlot(plotTypes, "bar") || !hasPlot(plotTypes, "line") {
		return nil, fmt.Errorf("unsupported mixed plot types")
	}

	assignMixedAxes(out.Series, plots)
	out.Plots = plotsToMixed(plots)
	out.AxisGroups = buildAxisGroups(axes)

	return out, nil
}

type plotState struct {
	plotType      string
	axisIDs       map[string]struct{}
	seriesIndices []int
	seriesCount   int
}

type axisInfo struct {
	kind              string
	id                string
	cross             string
	axisPos           string
	hasMajorGridlines bool
	hasMinorGridlines bool
}

func newPlotState(plotType string) *plotState {
	return &plotState{
		plotType: plotType,
		axisIDs:  make(map[string]struct{}),
	}
}

func assignMixedAxes(series []MixedSeries, plots []*plotState) {
	if len(plots) == 0 || len(series) == 0 {
		return
	}

	primaryIDs := plots[0].axisIDs
	plotAxis := make([]string, len(plots))
	for i, plot := range plots {
		plotAxis[i] = "primary"
		if len(primaryIDs) == 0 || len(plot.axisIDs) == 0 {
			continue
		}
		if !equalStringSets(primaryIDs, plot.axisIDs) {
			plotAxis[i] = "secondary"
		}
	}

	for i := range plots {
		for _, idx := range plots[i].seriesIndices {
			if idx >= 0 && idx < len(series) {
				series[idx].Axis = plotAxis[i]
			}
		}
	}
}

func equalStringSets(a, b map[string]struct{}) bool {
	if len(a) != len(b) {
		return false
	}
	for key := range a {
		if _, ok := b[key]; !ok {
			return false
		}
	}
	return true
}

func hasPlot(plotTypes map[string]struct{}, plot string) bool {
	_, ok := plotTypes[plot]
	return ok
}

func plotsToMixed(plots []*plotState) []MixedPlot {
	out := make([]MixedPlot, len(plots))
	for i, plot := range plots {
		out[i] = MixedPlot{
			PlotType:      plot.plotType,
			AxisIDs:       sortedKeys(plot.axisIDs),
			SeriesIndices: append([]int(nil), plot.seriesIndices...),
		}
	}
	return out
}

func sortedKeys(values map[string]struct{}) []string {
	keys := make([]string, 0, len(values))
	for key := range values {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

func buildAxisGroups(axes []axisInfo) []AxisGroup {
	catAxes := make(map[string]axisInfo)
	valAxes := make(map[string]axisInfo)
	for _, axis := range axes {
		if axis.id == "" {
			continue
		}
		if axis.kind == "cat" {
			catAxes[axis.id] = axis
		} else if axis.kind == "val" {
			valAxes[axis.id] = axis
		}
	}

	out := make([]AxisGroup, 0)
	seen := make(map[string]struct{})
	for _, cat := range catAxes {
		if cat.cross == "" {
			continue
		}
		val, ok := valAxes[cat.cross]
		if !ok || val.cross != cat.id {
			continue
		}
		key := cat.id + "|" + val.id
		if _, ok := seen[key]; ok {
			continue
		}
		seen[key] = struct{}{}
		out = append(out, AxisGroup{
			CatAxID:              cat.id,
			ValAxID:              val.id,
			ValAxisPos:           val.axisPos,
			ValHasMajorGridlines: val.hasMajorGridlines,
			ValHasMinorGridlines: val.hasMinorGridlines,
		})
	}

	sort.Slice(out, func(i, j int) bool {
		if out[i].CatAxID == out[j].CatAxID {
			return out[i].ValAxID < out[j].ValAxID
		}
		return out[i].CatAxID < out[j].CatAxID
	})
	return out
}

func isUnsupportedPlot(name string) bool {
	if name == "barChart" || name == "lineChart" {
		return false
	}
	if strings.HasSuffix(name, "Chart") {
		return true
	}
	return false
}
