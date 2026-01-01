package overlaystage

import (
	"fmt"

	"why-pptx/internal/ooxmlpkg"
)

type PackageOverlay struct {
	pkg      *ooxmlpkg.Package
	baseline map[string]struct{}
}

func NewPackageOverlay(pkg *ooxmlpkg.Package) (*PackageOverlay, error) {
	if pkg == nil {
		return nil, fmt.Errorf("package is nil")
	}
	parts, err := pkg.ListParts()
	if err != nil {
		return nil, err
	}

	baseline := make(map[string]struct{}, len(parts))
	for _, name := range parts {
		baseline[name] = struct{}{}
	}

	return &PackageOverlay{
		pkg:      pkg,
		baseline: baseline,
	}, nil
}

func (o *PackageOverlay) Get(path string) ([]byte, error) {
	if o == nil || o.pkg == nil {
		return nil, fmt.Errorf("overlay not initialized")
	}
	return o.pkg.ReadPart(path)
}

func (o *PackageOverlay) Set(path string, content []byte) error {
	if o == nil || o.pkg == nil {
		return fmt.Errorf("overlay not initialized")
	}
	o.pkg.WritePart(path, content)
	return nil
}

func (o *PackageOverlay) Has(path string) (bool, error) {
	if o == nil || o.pkg == nil {
		return false, fmt.Errorf("overlay not initialized")
	}

	if _, ok := o.baseline[path]; ok {
		return true, nil
	}

	parts, err := o.pkg.ListParts()
	if err != nil {
		return false, err
	}
	for _, name := range parts {
		if name == path {
			return true, nil
		}
	}
	return false, nil
}

func (o *PackageOverlay) ListEntries() ([]string, error) {
	if o == nil || o.pkg == nil {
		return nil, fmt.Errorf("overlay not initialized")
	}
	return o.pkg.ListParts()
}

func (o *PackageOverlay) HasBaseline(path string) (bool, error) {
	if o == nil || o.pkg == nil {
		return false, fmt.Errorf("overlay not initialized")
	}
	_, ok := o.baseline[path]
	return ok, nil
}
