package ooxmlpkg

import "errors"

var (
	ErrOpenFailed   = errors.New("ooxmlpkg: open failed")
	ErrPartNotFound = errors.New("ooxmlpkg: part not found")
	ErrSaveFailed   = errors.New("ooxmlpkg: save failed")
)
