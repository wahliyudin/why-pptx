package rels

import (
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

type Relationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string
}

type Rels struct {
	ByID map[string]Relationship
}

func Parse(r io.Reader) (*Rels, error) {
	decoder := xml.NewDecoder(r)
	out := &Rels{ByID: map[string]Relationship{}}

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse rels: %w", err)
		}

		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}

		rel := Relationship{}
		for _, attr := range start.Attr {
			switch strings.ToLower(attr.Name.Local) {
			case "id":
				rel.ID = attr.Value
			case "type":
				rel.Type = attr.Value
			case "target":
				rel.Target = attr.Value
			case "targetmode":
				rel.TargetMode = attr.Value
			}
		}

		if rel.ID != "" {
			out.ByID[rel.ID] = rel
		}
	}

	return out, nil
}

func (r *Rels) Resolve(id string) (Relationship, bool) {
	if r == nil || r.ByID == nil {
		return Relationship{}, false
	}
	rel, ok := r.ByID[id]
	return rel, ok
}
