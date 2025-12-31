package pptx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"path"
	"strings"

	"why-pptx/internal/chartxml"
	"why-pptx/internal/rels"
)

type ChartInfo struct {
	Index        int
	SlidePath    string
	ChartPath    string
	WorkbookPath string
	ChartType    string
	Title        string
	AltText      string
	SeriesCount  int
}

func (d *Document) ListCharts() ([]ChartInfo, error) {
	if d == nil || d.pkg == nil {
		return nil, fmt.Errorf("document not initialized")
	}

	charts, err := d.DiscoverEmbeddedCharts()
	if err != nil {
		return nil, err
	}
	if len(charts) == 0 {
		return []ChartInfo{}, nil
	}

	out := make([]ChartInfo, 0, len(charts))
	for i, chart := range charts {
		info := ChartInfo{
			Index:        i,
			SlidePath:    chart.SlidePath,
			ChartPath:    chart.ChartPath,
			WorkbookPath: chart.WorkbookPath,
			ChartType:    "unknown",
		}

		titleFromSlide, altText := d.slideChartAltText(chart.SlidePath, chart.ChartPath)
		info.AltText = altText

		data, err := d.pkg.ReadPart(chart.ChartPath)
		if err != nil {
			if err := d.handleChartInfoError(chart, fmt.Errorf("read chart %q: %w", chart.ChartPath, err)); err != nil {
				return nil, err
			}
			if info.Title == "" && titleFromSlide != "" {
				info.Title = titleFromSlide
			}
			out = append(out, info)
			continue
		}

		parsed, err := chartxml.ParseInfo(bytes.NewReader(data))
		if err != nil {
			if err := d.handleChartInfoError(chart, err); err != nil {
				return nil, err
			}
			if info.Title == "" && titleFromSlide != "" {
				info.Title = titleFromSlide
			}
			out = append(out, info)
			continue
		}

		info.ChartType = parsed.ChartType
		info.SeriesCount = parsed.SeriesCount
		info.Title = parsed.Title
		if info.Title == "" && titleFromSlide != "" {
			info.Title = titleFromSlide
		}

		out = append(out, info)
	}

	return out, nil
}

func (d *Document) ApplyChartDataByName(name string, data map[string][]string) error {
	if name == "" {
		return fmt.Errorf("chart name is required")
	}

	charts, err := d.ListCharts()
	if err != nil {
		return err
	}

	matches := matchChartsByName(charts, name)
	if len(matches) == 0 {
		return fmt.Errorf("chart not found")
	}
	if len(matches) > 1 {
		return d.handleChartNameAmbiguous(name, len(matches))
	}

	return d.ApplyChartData(matches[0].Index, data)
}

func (d *Document) ApplyChartDataByPath(chartPath string, data map[string][]string) error {
	if chartPath == "" {
		return fmt.Errorf("chart path is required")
	}

	charts, err := d.ListCharts()
	if err != nil {
		return err
	}

	for _, chart := range charts {
		if chart.ChartPath == chartPath {
			return d.ApplyChartData(chart.Index, data)
		}
	}

	return fmt.Errorf("chart not found")
}

func matchChartsByName(charts []ChartInfo, name string) []ChartInfo {
	matches := make([]ChartInfo, 0)
	for _, chart := range charts {
		if chart.Title != "" && chart.Title == name {
			matches = append(matches, chart)
		}
	}
	if len(matches) > 0 {
		return matches
	}
	for _, chart := range charts {
		if chart.AltText != "" && chart.AltText == name {
			matches = append(matches, chart)
		}
	}
	if len(matches) > 0 {
		return matches
	}
	for _, chart := range charts {
		if chart.Title != "" && strings.EqualFold(chart.Title, name) {
			matches = append(matches, chart)
			continue
		}
		if chart.AltText != "" && strings.EqualFold(chart.AltText, name) {
			matches = append(matches, chart)
		}
	}
	return matches
}

func (d *Document) handleChartInfoError(chart EmbeddedChart, err error) error {
	if d.opts.Mode != BestEffort {
		return err
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    "CHART_INFO_PARSE_FAILED",
		Message: "Failed to parse chart info; chart metadata is partial",
		Context: map[string]string{
			"slide": chart.SlidePath,
			"chart": chart.ChartPath,
			"error": err.Error(),
		},
	})

	return nil
}

func (d *Document) handleChartNameAmbiguous(name string, matches int) error {
	err := fmt.Errorf("ambiguous chart name")
	if d.opts.Mode != BestEffort {
		return err
	}

	d.addAlert(Alert{
		Level:   "warn",
		Code:    "CHART_NAME_AMBIGUOUS",
		Message: "Chart name is ambiguous; no chart selected",
		Context: map[string]string{
			"name":    name,
			"matches": fmt.Sprintf("%d", matches),
		},
	})

	return err
}

func (d *Document) slideChartAltText(slidePath, chartPath string) (string, string) {
	relID, err := d.findChartRelID(slidePath, chartPath)
	if err != nil || relID == "" {
		return "", ""
	}

	data, err := d.pkg.ReadPart(slidePath)
	if err != nil {
		return "", ""
	}

	title, descr, err := parseSlideChartProps(data, relID)
	if err != nil {
		return "", ""
	}
	return title, descr
}

func (d *Document) findChartRelID(slidePath, chartPath string) (string, error) {
	relsPath := slideRelsPath(slidePath)
	data, err := d.pkg.ReadPart(relsPath)
	if err != nil {
		return "", err
	}

	parsed, err := rels.Parse(bytes.NewReader(data))
	if err != nil {
		return "", err
	}

	for id, rel := range parsed.ByID {
		if !strings.HasSuffix(rel.Type, "/chart") {
			continue
		}
		if rel.TargetMode == "External" {
			continue
		}
		target := rels.ResolveTarget(slidePath, rel.Target)
		if target == chartPath {
			return id, nil
		}
	}

	return "", nil
}

func slideRelsPath(slidePath string) string {
	return path.Join(path.Dir(slidePath), "_rels", path.Base(slidePath)+".rels")
}

func parseSlideChartProps(data []byte, relID string) (string, string, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	inGraphicFrame := false
	graphicDepth := 0
	currentName := ""
	currentDescr := ""

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", "", fmt.Errorf("parse slide xml: %w", err)
		}

		switch tok := token.(type) {
		case xml.StartElement:
			if tok.Name.Local == "graphicFrame" {
				if !inGraphicFrame {
					inGraphicFrame = true
					graphicDepth = 1
					currentName = ""
					currentDescr = ""
				} else {
					graphicDepth++
				}
				continue
			}
			if inGraphicFrame {
				graphicDepth++
				if tok.Name.Local == "cNvPr" {
					for _, attr := range tok.Attr {
						switch attr.Name.Local {
						case "name":
							currentName = attr.Value
						case "descr":
							currentDescr = attr.Value
						}
					}
				}
				if tok.Name.Local == "chart" {
					chartRel := ""
					for _, attr := range tok.Attr {
						if attr.Name.Local == "id" {
							chartRel = attr.Value
							break
						}
					}
					if chartRel == relID {
						return currentName, currentDescr, nil
					}
				}
			}
		case xml.EndElement:
			if inGraphicFrame {
				graphicDepth--
				if graphicDepth == 0 {
					inGraphicFrame = false
					currentName = ""
					currentDescr = ""
				}
			}
		}
	}

	return "", "", nil
}
