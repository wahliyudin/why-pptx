package chartcache

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
)

type RangeKind string

const (
	KindCategories RangeKind = "categories"
	KindValues     RangeKind = "values"
	KindSeriesName RangeKind = "seriesName"
)

type Range struct {
	Kind        RangeKind
	SeriesIndex int
	Sheet       string
	StartCell   string
	EndCell     string
}

type Dependencies struct {
	ChartType string
	Ranges    []Range
}

type ValueProvider func(kind RangeKind, sheet, start, end string) ([]string, error)

func SyncCaches(chartXML []byte, deps Dependencies, provider ValueProvider) ([]byte, error) {
	if deps.ChartType != "bar" && deps.ChartType != "line" {
		return nil, fmt.Errorf("unsupported chart type %q", deps.ChartType)
	}

	seriesData, err := buildSeriesData(deps, provider)
	if err != nil {
		return nil, err
	}

	targetChart := "barChart"
	if deps.ChartType == "line" {
		targetChart = "lineChart"
	}

	decoder := xml.NewDecoder(bytes.NewReader(chartXML))
	var buf bytes.Buffer
	encoder := xml.NewEncoder(&buf)

	foundTarget := false
	inTarget := false
	var targetName xml.Name
	chartNS := ""

	currentSeries := -1
	seriesSeen := make(map[int]bool)

	catDepth := 0
	valDepth := 0
	txDepth := 0

	inRef := false
	refKind := RangeKind("")
	refName := xml.Name{}
	refHasCache := false

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
			if tok.Name.Local == targetChart && !foundTarget {
				foundTarget = true
				inTarget = true
				targetName = tok.Name
				chartNS = tok.Name.Space
			}
			if inTarget && tok.Name.Local == "ser" {
				currentSeries++
				seriesSeen[currentSeries] = true
			}
			if inTarget && currentSeries >= 0 {
				switch tok.Name.Local {
				case "cat":
					catDepth++
				case "val":
					valDepth++
				case "tx":
					txDepth++
				}

				if tok.Name.Local == "strRef" || tok.Name.Local == "numRef" {
					kind := refKindFor(tok.Name.Local, catDepth, valDepth, txDepth)
					if kind != "" && seriesHasData(seriesData, currentSeries, kind) {
						markRefSeen(seriesData, currentSeries, kind)
						inRef = true
						refKind = kind
						refName = tok.Name
						refHasCache = false
					}
				}
			}

			if inTarget && inRef && (tok.Name.Local == "strCache" || tok.Name.Local == "numCache") {
				if cacheMatchesRef(refKind, tok.Name.Local) {
					values := seriesValues(seriesData, currentSeries, refKind)
					if err := writeCache(encoder, tok.Name, tok.Attr, values, chartNS); err != nil {
						return nil, err
					}
					refHasCache = true
					markCacheUpdated(seriesData, currentSeries, refKind)
					if err := skipElement(decoder); err != nil {
						return nil, err
					}
					continue
				}
			}

			if err := encoder.EncodeToken(tok); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if inTarget && currentSeries >= 0 {
				switch tok.Name.Local {
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
				}

				if inRef && tok.Name.Local == refName.Local {
					if !refHasCache {
						values := seriesValues(seriesData, currentSeries, refKind)
						if len(values) > 0 || seriesHasData(seriesData, currentSeries, refKind) {
							cacheName := cacheNameFor(refKind, chartNS)
							if err := writeCache(encoder, cacheName, nil, values, chartNS); err != nil {
								return nil, err
							}
							markCacheUpdated(seriesData, currentSeries, refKind)
						}
					}
					inRef = false
					refKind = ""
					refName = xml.Name{}
					refHasCache = false
				}
			}

			if inTarget && tok.Name.Local == "ser" {
				if err := validateSeries(seriesData, currentSeries); err != nil {
					return nil, err
				}
			}

			if inTarget && tok.Name.Local == targetName.Local {
				inTarget = false
			}

			if err := encoder.EncodeToken(tok); err != nil {
				return nil, err
			}
		default:
			if err := encoder.EncodeToken(tok); err != nil {
				return nil, err
			}
		}
	}

	if err := encoder.Flush(); err != nil {
		return nil, err
	}
	if !foundTarget {
		return nil, fmt.Errorf("chart type %q not found", deps.ChartType)
	}

	for index := range seriesData {
		if !seriesSeen[index] {
			return nil, fmt.Errorf("chart series %d not found", index)
		}
	}

	return buf.Bytes(), nil
}

type seriesCache struct {
	categories []string
	values     []string
	name       []string

	catRefSeen bool
	valRefSeen bool
	txRefSeen  bool

	catUpdated bool
	valUpdated bool
	txUpdated  bool
}

func buildSeriesData(deps Dependencies, provider ValueProvider) (map[int]*seriesCache, error) {
	series := make(map[int]*seriesCache)

	for _, r := range deps.Ranges {
		if r.SeriesIndex < 0 {
			continue
		}
		values, err := provider(r.Kind, r.Sheet, r.StartCell, r.EndCell)
		if err != nil {
			return nil, err
		}
		entry := series[r.SeriesIndex]
		if entry == nil {
			entry = &seriesCache{}
			series[r.SeriesIndex] = entry
		}

		switch r.Kind {
		case KindCategories:
			if entry.categories != nil {
				return nil, fmt.Errorf("duplicate categories range for series %d", r.SeriesIndex)
			}
			entry.categories = values
		case KindValues:
			if entry.values != nil {
				return nil, fmt.Errorf("duplicate values range for series %d", r.SeriesIndex)
			}
			entry.values = values
		case KindSeriesName:
			if entry.name != nil {
				return nil, fmt.Errorf("duplicate series name range for series %d", r.SeriesIndex)
			}
			entry.name = values
		default:
			return nil, fmt.Errorf("unsupported range kind %q", r.Kind)
		}
	}

	return series, nil
}

func seriesHasData(series map[int]*seriesCache, index int, kind RangeKind) bool {
	entry := series[index]
	if entry == nil {
		return false
	}
	switch kind {
	case KindCategories:
		return entry.categories != nil
	case KindValues:
		return entry.values != nil
	case KindSeriesName:
		return entry.name != nil
	default:
		return false
	}
}

func seriesValues(series map[int]*seriesCache, index int, kind RangeKind) []string {
	entry := series[index]
	if entry == nil {
		return nil
	}
	switch kind {
	case KindCategories:
		return entry.categories
	case KindValues:
		return entry.values
	case KindSeriesName:
		return entry.name
	default:
		return nil
	}
}

func markRefSeen(series map[int]*seriesCache, index int, kind RangeKind) {
	entry := series[index]
	if entry == nil {
		return
	}
	switch kind {
	case KindCategories:
		entry.catRefSeen = true
	case KindValues:
		entry.valRefSeen = true
	case KindSeriesName:
		entry.txRefSeen = true
	}
}

func markCacheUpdated(series map[int]*seriesCache, index int, kind RangeKind) {
	entry := series[index]
	if entry == nil {
		return
	}
	switch kind {
	case KindCategories:
		entry.catUpdated = true
	case KindValues:
		entry.valUpdated = true
	case KindSeriesName:
		entry.txUpdated = true
	}
}

func validateSeries(series map[int]*seriesCache, index int) error {
	entry := series[index]
	if entry == nil {
		return nil
	}
	if entry.categories != nil && !entry.catRefSeen {
		return fmt.Errorf("missing category reference for series %d", index)
	}
	if entry.values != nil && !entry.valRefSeen {
		return fmt.Errorf("missing values reference for series %d", index)
	}
	if entry.name != nil && !entry.txRefSeen {
		return fmt.Errorf("missing series name reference for series %d", index)
	}
	return nil
}

func refKindFor(refName string, catDepth, valDepth, txDepth int) RangeKind {
	switch refName {
	case "strRef":
		if catDepth > 0 {
			return KindCategories
		}
		if txDepth > 0 {
			return KindSeriesName
		}
	case "numRef":
		if valDepth > 0 {
			return KindValues
		}
	}
	return ""
}

func cacheMatchesRef(kind RangeKind, cacheName string) bool {
	switch kind {
	case KindCategories, KindSeriesName:
		return cacheName == "strCache"
	case KindValues:
		return cacheName == "numCache"
	default:
		return false
	}
}

func cacheNameFor(kind RangeKind, space string) xml.Name {
	local := "strCache"
	if kind == KindValues {
		local = "numCache"
	}
	return xml.Name{Space: space, Local: local}
}

func writeCache(encoder *xml.Encoder, name xml.Name, attrs []xml.Attr, values []string, space string) error {
	start := xml.StartElement{Name: name, Attr: attrs}
	if err := encoder.EncodeToken(start); err != nil {
		return err
	}

	countAttr := xml.Attr{Name: xml.Name{Local: "val"}, Value: fmt.Sprintf("%d", len(values))}
	ptCount := xml.StartElement{Name: xml.Name{Space: space, Local: "ptCount"}, Attr: []xml.Attr{countAttr}}
	if err := encoder.EncodeToken(ptCount); err != nil {
		return err
	}
	if err := encoder.EncodeToken(xml.EndElement{Name: ptCount.Name}); err != nil {
		return err
	}

	for i, val := range values {
		pt := xml.StartElement{
			Name: xml.Name{Space: space, Local: "pt"},
			Attr: []xml.Attr{{Name: xml.Name{Local: "idx"}, Value: fmt.Sprintf("%d", i)}},
		}
		if err := encoder.EncodeToken(pt); err != nil {
			return err
		}
		v := xml.StartElement{Name: xml.Name{Space: space, Local: "v"}}
		if err := encoder.EncodeToken(v); err != nil {
			return err
		}
		if val != "" {
			if err := encoder.EncodeToken(xml.CharData([]byte(val))); err != nil {
				return err
			}
		}
		if err := encoder.EncodeToken(xml.EndElement{Name: v.Name}); err != nil {
			return err
		}
		if err := encoder.EncodeToken(xml.EndElement{Name: pt.Name}); err != nil {
			return err
		}
	}

	if err := encoder.EncodeToken(xml.EndElement{Name: name}); err != nil {
		return err
	}
	return nil
}

func skipElement(decoder *xml.Decoder) error {
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			return err
		}
		switch token.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
	}
	return nil
}
