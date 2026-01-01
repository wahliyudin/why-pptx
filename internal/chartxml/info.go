package chartxml

import (
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

type Info struct {
	ChartType   string
	SeriesCount int
	Title       string
}

func ParseInfo(r io.Reader) (*Info, error) {
	decoder := xml.NewDecoder(r)
	info := &Info{ChartType: "unknown"}

	barDepth := 0
	lineDepth := 0
	pieDepth := 0
	areaDepth := 0
	titleDepth := 0
	inTitleText := false
	titleSet := false
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
				info.ChartType = updateChartType(info.ChartType, "bar")
			case "lineChart":
				lineDepth++
				info.ChartType = updateChartType(info.ChartType, "line")
			case "pieChart":
				pieDepth++
				info.ChartType = updateChartType(info.ChartType, "pie")
			case "areaChart":
				areaDepth++
				info.ChartType = updateChartType(info.ChartType, "area")
			case "ser":
				if barDepth > 0 || lineDepth > 0 || pieDepth > 0 || areaDepth > 0 {
					info.SeriesCount++
				}
			case "title":
				titleDepth++
			case "t", "v":
				if titleDepth > 0 && !titleSet {
					inTitleText = true
					buf.Reset()
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
			case "title":
				if titleDepth > 0 {
					titleDepth--
				}
			case "t", "v":
				if inTitleText {
					text := strings.TrimSpace(buf.String())
					if text != "" && !titleSet {
						info.Title = text
						titleSet = true
					}
					inTitleText = false
					buf.Reset()
				}
			}
		case xml.CharData:
			if inTitleText {
				buf.Write([]byte(tok))
			}
		}
	}

	return info, nil
}
