package ooxmlpkg

import (
	"archive/zip"
	"bytes"
	"compress/flate"
	"fmt"
	"hash/crc32"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

type Package struct {
	data    []byte
	reader  *zip.Reader
	index   map[string]*zip.File
	overlay map[string][]byte
}

func OpenFile(path string) (*Package, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("%w: %s: %v", ErrOpenFailed, path, err)
	}

	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("%w: %s: %v", ErrOpenFailed, path, err)
	}

	index := make(map[string]*zip.File, len(reader.File))
	for _, part := range reader.File {
		index[part.Name] = part
	}

	return &Package{
		data:    data,
		reader:  reader,
		index:   index,
		overlay: make(map[string][]byte),
	}, nil
}

func (p *Package) ListParts() ([]string, error) {
	if p == nil || p.reader == nil {
		return nil, fmt.Errorf("%w: package not initialized", ErrOpenFailed)
	}

	names := make([]string, 0, len(p.reader.File)+len(p.overlay))
	seen := make(map[string]struct{}, len(p.reader.File)+len(p.overlay))
	for _, part := range p.reader.File {
		names = append(names, part.Name)
		seen[part.Name] = struct{}{}
	}

	if len(p.overlay) == 0 {
		return names, nil
	}

	extras := make([]string, 0, len(p.overlay))
	for name := range p.overlay {
		if _, ok := seen[name]; !ok {
			extras = append(extras, name)
		}
	}
	sort.Strings(extras)
	names = append(names, extras...)

	return names, nil
}

func (p *Package) ReadPart(name string) ([]byte, error) {
	if p == nil {
		return nil, fmt.Errorf("%w: package not initialized", ErrOpenFailed)
	}

	if data, ok := p.overlay[name]; ok {
		return append([]byte(nil), data...), nil
	}

	part, ok := p.index[name]
	if !ok {
		return nil, fmt.Errorf("%w: %s", ErrPartNotFound, name)
	}

	reader, err := part.Open()
	if err != nil {
		return nil, fmt.Errorf("read part %q: %w", name, err)
	}
	defer reader.Close()

	data, err := io.ReadAll(reader)
	if err != nil {
		return nil, fmt.Errorf("read part %q: %w", name, err)
	}

	return data, nil
}

func (p *Package) WritePart(name string, data []byte) {
	if p == nil {
		return
	}

	if p.overlay == nil {
		p.overlay = make(map[string][]byte)
	}

	copied := make([]byte, len(data))
	copy(copied, data)
	p.overlay[name] = copied
}

func (p *Package) SaveFile(path string) error {
	if p == nil || p.reader == nil {
		return fmt.Errorf("%w: package not initialized", ErrSaveFailed)
	}

	dir := filepath.Dir(path)
	base := filepath.Base(path)
	tmpFile, err := os.CreateTemp(dir, base+".tmp-*")
	if err != nil {
		return fmt.Errorf("%w: %s: %v", ErrSaveFailed, path, err)
	}
	tmpName := tmpFile.Name()
	cleanup := true
	defer func() {
		if cleanup {
			_ = os.Remove(tmpName)
		}
	}()

	writer := zip.NewWriter(tmpFile)
	written := make(map[string]struct{}, len(p.reader.File)+len(p.overlay))

	for _, part := range p.reader.File {
		name := part.Name
		if data, ok := p.overlay[name]; ok {
			if err := writeOverrideEntry(writer, part, data); err != nil {
				_ = writer.Close()
				_ = tmpFile.Close()
				return fmt.Errorf("%w: write part %q: %v", ErrSaveFailed, name, err)
			}
		} else {
			if err := writer.Copy(part); err != nil {
				_ = writer.Close()
				_ = tmpFile.Close()
				return fmt.Errorf("%w: copy part %q: %v", ErrSaveFailed, name, err)
			}
		}
		written[name] = struct{}{}
	}

	for name, data := range p.overlay {
		if _, ok := written[name]; ok {
			continue
		}
		if err := writeNewEntry(writer, name, data); err != nil {
			_ = writer.Close()
			_ = tmpFile.Close()
			return fmt.Errorf("%w: write part %q: %v", ErrSaveFailed, name, err)
		}
	}

	if err := writer.Close(); err != nil {
		_ = tmpFile.Close()
		return fmt.Errorf("%w: %s: %v", ErrSaveFailed, path, err)
	}

	if err := tmpFile.Sync(); err != nil {
		_ = tmpFile.Close()
		return fmt.Errorf("%w: %s: %v", ErrSaveFailed, path, err)
	}
	if err := tmpFile.Close(); err != nil {
		return fmt.Errorf("%w: %s: %v", ErrSaveFailed, path, err)
	}

	if err := replaceFile(tmpName, path); err != nil {
		return fmt.Errorf("%w: %s: %v", ErrSaveFailed, path, err)
	}

	cleanup = false
	return nil
}

func writeNewEntry(writer *zip.Writer, name string, data []byte) error {
	if strings.HasSuffix(name, "/") {
		header := zip.FileHeader{Name: name, Method: zip.Store}
		return writeDirectoryEntry(writer, &header)
	}

	header := zip.FileHeader{Name: name, Method: zip.Deflate}
	return writeRawEntry(writer, &header, data)
}

func replaceFile(src, dst string) error {
	if err := os.Rename(src, dst); err == nil {
		return nil
	} else {
		if _, statErr := os.Stat(dst); statErr == nil {
			if removeErr := os.Remove(dst); removeErr != nil {
				return removeErr
			}
			return os.Rename(src, dst)
		}
		return err
	}
}

func writeOverrideEntry(writer *zip.Writer, part *zip.File, data []byte) error {
	if part.FileInfo().IsDir() {
		return writer.Copy(part)
	}

	header := part.FileHeader
	if header.Flags&0x8 != 0 {
		return writeEntryWithDescriptor(writer, &header, data)
	}

	return writeRawEntry(writer, &header, data)
}

func writeEntryWithDescriptor(writer *zip.Writer, header *zip.FileHeader, data []byte) error {
	header.CRC32 = 0
	header.CompressedSize = 0
	header.UncompressedSize = 0
	header.CompressedSize64 = 0
	header.UncompressedSize64 = 0

	entry, err := writer.CreateHeader(header)
	if err != nil {
		return err
	}

	_, err = entry.Write(data)
	return err
}

func writeRawEntry(writer *zip.Writer, header *zip.FileHeader, data []byte) error {
	// Precompute sizes/CRC and use CreateRaw so we don't introduce data descriptors.
	compressed, err := compressData(header.Method, data)
	if err != nil {
		return err
	}

	header.Flags &^= 0x8
	header.CRC32 = crc32.ChecksumIEEE(data)
	header.UncompressedSize64 = uint64(len(data))
	header.UncompressedSize = uint32(len(data))
	header.CompressedSize64 = uint64(len(compressed))
	header.CompressedSize = uint32(len(compressed))

	entry, err := writer.CreateRaw(header)
	if err != nil {
		return err
	}
	if len(compressed) == 0 {
		return nil
	}

	_, err = entry.Write(compressed)
	return err
}

func writeDirectoryEntry(writer *zip.Writer, header *zip.FileHeader) error {
	header.Method = zip.Store
	header.Flags &^= 0x8
	header.CRC32 = 0
	header.UncompressedSize64 = 0
	header.UncompressedSize = 0
	header.CompressedSize64 = 0
	header.CompressedSize = 0

	_, err := writer.CreateHeader(header)
	return err
}

func compressData(method uint16, data []byte) ([]byte, error) {
	switch method {
	case zip.Store:
		return data, nil
	case zip.Deflate:
		var buf bytes.Buffer
		zw, err := flate.NewWriter(&buf, flate.DefaultCompression)
		if err != nil {
			return nil, err
		}
		if len(data) > 0 {
			if _, err := zw.Write(data); err != nil {
				_ = zw.Close()
				return nil, err
			}
		}
		if err := zw.Close(); err != nil {
			return nil, err
		}
		return buf.Bytes(), nil
	default:
		return nil, fmt.Errorf("unsupported compression method: %d", method)
	}
}
