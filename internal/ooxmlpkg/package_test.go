package ooxmlpkg

import (
	"archive/zip"
	"errors"
	"hash/crc32"
	"io"
	"os"
	"path/filepath"
	"testing"
)

func TestPackageReadWriteSave(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	original := map[string][]byte{
		"[Content_Types].xml":   []byte("types"),
		"ppt/presentation.xml":  []byte("original"),
		"ppt/slides/slide1.xml": []byte("slide1"),
	}

	if err := writeZip(inputPath, original); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	pkg, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	got, err := pkg.ReadPart("ppt/presentation.xml")
	if err != nil {
		t.Fatalf("ReadPart: %v", err)
	}
	if string(got) != "original" {
		t.Fatalf("ReadPart content mismatch: got %q", string(got))
	}

	pkg.WritePart("ppt/presentation.xml", []byte("updated"))
	pkg.WritePart("ppt/new.xml", []byte("new"))

	if err := pkg.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	outParts, err := readZip(outputPath)
	if err != nil {
		t.Fatalf("readZip: %v", err)
	}

	if string(outParts["ppt/presentation.xml"]) != "updated" {
		t.Fatalf("updated part mismatch: got %q", string(outParts["ppt/presentation.xml"]))
	}
	if string(outParts["ppt/slides/slide1.xml"]) != "slide1" {
		t.Fatalf("untouched part mismatch: got %q", string(outParts["ppt/slides/slide1.xml"]))
	}
	if string(outParts["ppt/new.xml"]) != "new" {
		t.Fatalf("new part mismatch: got %q", string(outParts["ppt/new.xml"]))
	}
	if string(outParts["[Content_Types].xml"]) != "types" {
		t.Fatalf("root part mismatch: got %q", string(outParts["[Content_Types].xml"]))
	}
}

func TestReadPartMissing(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")

	if err := writeZip(inputPath, map[string][]byte{
		"ppt/presentation.xml": []byte("original"),
	}); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	pkg, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	_, err = pkg.ReadPart("missing.xml")
	if !errors.Is(err, ErrPartNotFound) {
		t.Fatalf("expected ErrPartNotFound, got %v", err)
	}
}

func TestOpenFileMissing(t *testing.T) {
	dir := t.TempDir()
	missingPath := filepath.Join(dir, "missing.pptx")

	_, err := OpenFile(missingPath)
	if !errors.Is(err, ErrOpenFailed) {
		t.Fatalf("expected ErrOpenFailed, got %v", err)
	}
}

func TestSaveFileReplacesExisting(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	original := map[string][]byte{
		"ppt/presentation.xml": []byte("original"),
	}

	if err := writeZip(inputPath, original); err != nil {
		t.Fatalf("writeZip: %v", err)
	}

	if err := os.WriteFile(outputPath, []byte("stale"), 0o600); err != nil {
		t.Fatalf("WriteFile: %v", err)
	}

	pkg, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	if err := pkg.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	outParts, err := readZip(outputPath)
	if err != nil {
		t.Fatalf("readZip: %v", err)
	}
	if string(outParts["ppt/presentation.xml"]) != "original" {
		t.Fatalf("output part mismatch: got %q", string(outParts["ppt/presentation.xml"]))
	}
}

func TestSaveFileDoesNotAddDataDescriptorFlag(t *testing.T) {
	dir := t.TempDir()
	inputPath := filepath.Join(dir, "input.pptx")
	outputPath := filepath.Join(dir, "output.pptx")

	data := []byte("original")
	if err := writeZipWithHeader(inputPath, "ppt/presentation.xml", data, 0); err != nil {
		t.Fatalf("writeZipWithHeader: %v", err)
	}
	if flags, err := readZipFlags(inputPath, "ppt/presentation.xml"); err != nil {
		t.Fatalf("readZipFlags(input): %v", err)
	} else if flags&0x8 != 0 {
		t.Fatalf("input unexpectedly uses data descriptor: 0x%x", flags)
	}

	pkg, err := OpenFile(inputPath)
	if err != nil {
		t.Fatalf("OpenFile: %v", err)
	}

	pkg.WritePart("ppt/presentation.xml", []byte("updated"))
	if err := pkg.SaveFile(outputPath); err != nil {
		t.Fatalf("SaveFile: %v", err)
	}

	flags, err := readZipFlags(outputPath, "ppt/presentation.xml")
	if err != nil {
		t.Fatalf("readZipFlags: %v", err)
	}
	if flags&0x8 != 0 {
		t.Fatalf("unexpected data descriptor flag set: 0x%x", flags)
	}
}

func writeZip(path string, parts map[string][]byte) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := zip.NewWriter(file)
	for name, data := range parts {
		entry, err := writer.Create(name)
		if err != nil {
			_ = writer.Close()
			return err
		}
		if _, err := entry.Write(data); err != nil {
			_ = writer.Close()
			return err
		}
	}

	if err := writer.Close(); err != nil {
		return err
	}

	return nil
}

func writeZipWithHeader(path, name string, data []byte, flags uint16) error {
	file, err := os.Create(path)
	if err != nil {
		return err
	}
	defer file.Close()

	writer := zip.NewWriter(file)
	header := zip.FileHeader{
		Name:               name,
		Method:             zip.Store,
		Flags:              flags,
		CRC32:              crc32.ChecksumIEEE(data),
		UncompressedSize64: uint64(len(data)),
		CompressedSize64:   uint64(len(data)),
		UncompressedSize:   uint32(len(data)),
		CompressedSize:     uint32(len(data)),
	}
	entry, err := writer.CreateRaw(&header)
	if err != nil {
		_ = writer.Close()
		return err
	}
	if _, err := entry.Write(data); err != nil {
		_ = writer.Close()
		return err
	}
	if err := writer.Close(); err != nil {
		return err
	}

	return nil
}

func readZip(path string) (map[string][]byte, error) {
	file, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	info, err := file.Stat()
	if err != nil {
		return nil, err
	}

	reader, err := zip.NewReader(file, info.Size())
	if err != nil {
		return nil, err
	}

	out := make(map[string][]byte, len(reader.File))
	for _, part := range reader.File {
		rc, err := part.Open()
		if err != nil {
			return nil, err
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return nil, err
		}
		out[part.Name] = data
	}

	return out, nil
}

func readZipFlags(path, name string) (uint16, error) {
	file, err := os.Open(path)
	if err != nil {
		return 0, err
	}
	defer file.Close()

	info, err := file.Stat()
	if err != nil {
		return 0, err
	}

	reader, err := zip.NewReader(file, info.Size())
	if err != nil {
		return 0, err
	}

	for _, part := range reader.File {
		if part.Name == name {
			return part.Flags, nil
		}
	}

	return 0, os.ErrNotExist
}
