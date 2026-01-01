package chartxml

import (
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

const (
	KindCategories = "categories"
	KindValues     = "values"
	KindSeriesName = "seriesName"
)

type Formula struct {
	Kind        string
	SeriesIndex int
	Formula     string
}

type ParsedChart struct {
	ChartType string
	Formulas  []Formula
}

func Parse(r io.Reader) (*ParsedChart, error) {
	decoder := xml.NewDecoder(r)
	out := &ParsedChart{ChartType: "unknown"}

	seriesIndex := -1
	inSeries := false
	catDepth := 0
	valDepth := 0
	txDepth := 0
	barDepth := 0
	lineDepth := 0
	pieDepth := 0
	areaDepth := 0
	otherDepth := 0

	inFormula := false
	formulaKind := ""
	formulaSeries := -1
	var buf strings.Builder

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse chart xml: %w", err)
		}

		switch tok := token.(type) {
		case xml.StartElement:
			switch tok.Name.Local {
			case "barChart":
				barDepth++
				out.ChartType = updateChartType(out.ChartType, "bar")
			case "lineChart":
				lineDepth++
				out.ChartType = updateChartType(out.ChartType, "line")
			case "pieChart":
				pieDepth++
				out.ChartType = updateChartType(out.ChartType, "pie")
			case "areaChart":
				areaDepth++
				out.ChartType = updateChartType(out.ChartType, "area")
			default:
				if isOtherChart(tok.Name.Local) {
					otherDepth++
					out.ChartType = updateChartType(out.ChartType, "other")
				}
			case "ser":
				if barDepth+lineDepth+pieDepth+areaDepth > 0 {
					seriesIndex++
					inSeries = true
				}
			case "cat":
				if inSeries {
					catDepth++
				}
			case "val":
				if inSeries {
					valDepth++
				}
			case "tx":
				if inSeries {
					txDepth++
				}
			case "f":
				if inSeries {
					kind := ""
					if catDepth > 0 {
						kind = KindCategories
					} else if valDepth > 0 {
						kind = KindValues
					} else if txDepth > 0 {
						kind = KindSeriesName
					}
					if kind != "" {
						inFormula = true
						formulaKind = kind
						formulaSeries = seriesIndex
						buf.Reset()
					}
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
			case "pieChart":
				if pieDepth > 0 {
					pieDepth--
				}
			case "areaChart":
				if areaDepth > 0 {
					areaDepth--
				}
			default:
				if isOtherChart(tok.Name.Local) && otherDepth > 0 {
					otherDepth--
				}
			case "ser":
				inSeries = false
				catDepth = 0
				valDepth = 0
				txDepth = 0
				inFormula = false
				formulaKind = ""
				formulaSeries = -1
				buf.Reset()
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
				if inFormula {
					text := strings.TrimSpace(buf.String())
					if text != "" {
						out.Formulas = append(out.Formulas, Formula{
							Kind:        formulaKind,
							SeriesIndex: formulaSeries,
							Formula:     text,
						})
					}
					inFormula = false
					formulaKind = ""
					formulaSeries = -1
					buf.Reset()
				}
			}
		case xml.CharData:
			if inFormula {
				buf.Write([]byte(tok))
			}
		}
	}

	return out, nil
}

func updateChartType(current, next string) string {
	if current == "unknown" {
		return next
	}
	if current == next || current == "mixed" {
		return current
	}
	return "mixed"
}

func isOtherChart(name string) bool {
	if name == "barChart" || name == "lineChart" || name == "pieChart" || name == "areaChart" {
		return false
	}
	return strings.HasSuffix(name, "Chart")
}
