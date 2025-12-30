package chartdiscover

import (
	"bytes"
	"errors"
	"path"
	"sort"
	"strings"

	"why-pptx/internal/ooxmlpkg"
	"why-pptx/internal/rels"
)

type ChartRef struct {
	SlidePath string
	ChartPath string
}

type EmbeddedChart struct {
	SlidePath    string
	ChartPath    string
	WorkbookPath string
}

type SkippedChart struct {
	SlidePath string
	ChartPath string
	Reason    string
	Target    string
	RelsPath  string
}

type PartReader interface {
	ListParts() ([]string, error)
	ReadPart(name string) ([]byte, error)
}

func DiscoverChartRefs(pkg PartReader) ([]ChartRef, error) {
	parts, err := pkg.ListParts()
	if err != nil {
		return nil, err
	}

	slides := make([]string, 0)
	for _, part := range parts {
		match, _ := path.Match("ppt/slides/slide*.xml", part)
		if match {
			slides = append(slides, part)
		}
	}
	sort.Strings(slides)

	var refs []ChartRef
	for _, slide := range slides {
		relsPath := slideRelsPath(slide)
		data, err := pkg.ReadPart(relsPath)
		if err != nil {
			if errors.Is(err, ooxmlpkg.ErrPartNotFound) {
				continue
			}
			return nil, err
		}

		parsed, err := rels.Parse(bytes.NewReader(data))
		if err != nil {
			return nil, err
		}

		for _, rel := range parsed.ByID {
			if !strings.HasSuffix(rel.Type, "/chart") {
				continue
			}
			if rel.TargetMode == "External" {
				continue
			}
			target := rels.ResolveTarget(slide, rel.Target)
			refs = append(refs, ChartRef{
				SlidePath: slide,
				ChartPath: target,
			})
		}
	}

	return refs, nil
}

const (
	ReasonLinked           = "linked"
	ReasonRelsMissing      = "rels_missing"
	ReasonWorkbookNotFound = "workbook_not_found"
	ReasonUnsupported      = "unsupported_target"
)

func DiscoverEmbeddedCharts(pkg PartReader) ([]EmbeddedChart, []SkippedChart, error) {
	refs, err := DiscoverChartRefs(pkg)
	if err != nil {
		return nil, nil, err
	}

	embedded := make([]EmbeddedChart, 0, len(refs))
	skipped := make([]SkippedChart, 0)

	for _, ref := range refs {
		relsPath := chartRelsPath(ref.ChartPath)
		data, err := pkg.ReadPart(relsPath)
		if err != nil {
			if errors.Is(err, ooxmlpkg.ErrPartNotFound) {
				skipped = append(skipped, SkippedChart{
					SlidePath: ref.SlidePath,
					ChartPath: ref.ChartPath,
					Reason:    ReasonRelsMissing,
					RelsPath:  relsPath,
				})
				continue
			}
			return nil, nil, err
		}

		parsed, err := rels.Parse(bytes.NewReader(data))
		if err != nil {
			return nil, nil, err
		}

		embeddedPath := ""
		linkedTarget := ""
		unsupportedTarget := ""
		foundWorkbookRel := false
		for _, rel := range parsed.ByID {
			if !isWorkbookCandidate(rel) {
				continue
			}
			foundWorkbookRel = true
			if rel.TargetMode == "External" {
				if linkedTarget == "" {
					linkedTarget = rel.Target
				}
				continue
			}

			target := rels.ResolveTarget(ref.ChartPath, rel.Target)
			lowerTarget := strings.ToLower(target)
			if strings.HasPrefix(target, "ppt/embeddings/") && strings.HasSuffix(lowerTarget, ".xlsx") {
				if embeddedPath == "" {
					embeddedPath = target
				}
				continue
			}

			if unsupportedTarget == "" {
				unsupportedTarget = target
			}
		}

		if linkedTarget != "" {
			skipped = append(skipped, SkippedChart{
				SlidePath: ref.SlidePath,
				ChartPath: ref.ChartPath,
				Reason:    ReasonLinked,
				Target:    linkedTarget,
			})
			continue
		}
		if unsupportedTarget != "" {
			skipped = append(skipped, SkippedChart{
				SlidePath: ref.SlidePath,
				ChartPath: ref.ChartPath,
				Reason:    ReasonUnsupported,
				Target:    unsupportedTarget,
			})
			continue
		}
		if embeddedPath != "" {
			embedded = append(embedded, EmbeddedChart{
				SlidePath:    ref.SlidePath,
				ChartPath:    ref.ChartPath,
				WorkbookPath: embeddedPath,
			})
			continue
		}
		if !foundWorkbookRel {
			skipped = append(skipped, SkippedChart{
				SlidePath: ref.SlidePath,
				ChartPath: ref.ChartPath,
				Reason:    ReasonWorkbookNotFound,
			})
		}
	}

	return embedded, skipped, nil
}

func slideRelsPath(slidePath string) string {
	return path.Join(path.Dir(slidePath), "_rels", path.Base(slidePath)+".rels")
}

func chartRelsPath(chartPath string) string {
	return path.Join(path.Dir(chartPath), "_rels", path.Base(chartPath)+".rels")
}

func isWorkbookCandidate(rel rels.Relationship) bool {
	if rel.TargetMode == "External" {
		return true
	}
	if strings.HasSuffix(rel.Type, "/package") {
		return true
	}
	return strings.HasSuffix(strings.ToLower(rel.Target), ".xlsx")
}
