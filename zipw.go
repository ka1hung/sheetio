package sheetio

import (
	"archive/zip"
	"bytes"
	"compress/flate"
	"hash"
	"hash/crc32"
	"io"
	"time"
)

// cleanZip writes a "traditional" zip: every entry has the correct CRC and sizes in its local header,
// no data descriptor (flag 0x8), and valid DOS date and version fields.
// Go's zip.Writer.Create streams entries, so it always adds a data descriptor and leaves the date at 0;
// LibreOffice accepts that, but Excel is stricter when reading ODS.
// Each entry is compressed in memory first (only the compressed data is kept) and then written with CreateRaw.
type cleanZip struct {
	zw      *zip.Writer
	dosDate uint16
	dosTime uint16
}

func newCleanZip(w io.Writer, now time.Time) *cleanZip {
	if now.Year() < 1980 {
		now = time.Date(1980, 1, 1, 0, 0, 0, 0, time.UTC)
	}
	return &cleanZip{
		zw:      zip.NewWriter(w),
		dosDate: uint16((now.Year()-1980)<<9 | int(now.Month())<<5 | now.Day()),
		dosTime: uint16(now.Hour()<<11 | now.Minute()<<5 | now.Second()/2),
	}
}

func (z *cleanZip) header(name string, method uint16, crc uint32, csize, usize int64) *zip.FileHeader {
	return &zip.FileHeader{
		Name:               name,
		Method:             method,
		CRC32:              crc,
		CompressedSize64:   uint64(csize),
		UncompressedSize64: uint64(usize),
		CreatorVersion:     20,
		ReaderVersion:      20,
		ModifiedDate:       z.dosDate, //nolint:staticcheck // CreateRaw does not derive the DOS date from Modified
		ModifiedTime:       z.dosTime, //nolint:staticcheck
	}
}

// add writes a complete entry; fill produces the uncompressed content.
func (z *cleanZip) add(name string, method uint16, fill func(io.Writer) error) error {
	if method == zip.Store {
		var raw bytes.Buffer
		if err := fill(&raw); err != nil {
			return err
		}
		w, err := z.zw.CreateRaw(z.header(name, zip.Store, crc32.ChecksumIEEE(raw.Bytes()), int64(raw.Len()), int64(raw.Len())))
		if err != nil {
			return err
		}
		_, err = w.Write(raw.Bytes())
		return err
	}
	s := newSegment()
	if err := fill(s); err != nil {
		return err
	}
	if err := s.finish(true); err != nil {
		return err
	}
	return z.addSegments(name, []*segment{s})
}

// addSegments joins several separately compressed deflate segments into one entry.
// Every segment except the last ends with a sync flush (a non-final block), so they can be concatenated directly;
// the CRC is combined in final order with crc32Combine. This lets ODS write the rows first and the column definitions later
// without keeping the uncompressed content.
func (z *cleanZip) addSegments(name string, segs []*segment) error {
	var crc uint32
	var usize, csize int64
	for i, s := range segs {
		if i == 0 {
			crc = s.crc
		} else {
			crc = crc32Combine(crc, s.crc, s.n)
		}
		usize += s.n
		csize += int64(s.buf.Len())
	}
	w, err := z.zw.CreateRaw(z.header(name, zip.Deflate, crc, csize, usize))
	if err != nil {
		return err
	}
	for _, s := range segs {
		if _, err := w.Write(s.buf.Bytes()); err != nil {
			return err
		}
		s.buf = bytes.Buffer{} // release
	}
	return nil
}

func (z *cleanZip) close() error { return z.zw.Close() }

// segment is an independent deflate stream.
type segment struct {
	buf  bytes.Buffer
	fw   *flate.Writer
	h    hash.Hash32
	crc  uint32
	n    int64
	done bool
}

func newSegment() *segment {
	s := &segment{h: crc32.NewIEEE()}
	s.fw, _ = flate.NewWriter(&s.buf, flate.DefaultCompression) // fails only for an invalid level
	return s
}

func (s *segment) Write(p []byte) (int, error) {
	s.h.Write(p)
	s.n += int64(len(p))
	return s.fw.Write(p)
}

// finish ends the segment; final=true writes the final block, otherwise only a sync flush is done.
func (s *segment) finish(final bool) error {
	if s.done {
		return nil
	}
	s.done = true
	s.crc = s.h.Sum32()
	if final {
		return s.fw.Close()
	}
	return s.fw.Flush()
}

// crc32Combine computes crc(A||B); the algorithm is taken from zlib's crc32_combine.
func crc32Combine(crc1, crc2 uint32, len2 int64) uint32 {
	if len2 <= 0 {
		return crc1
	}
	var even, odd [32]uint32
	odd[0] = 0xedb88320 // IEEE polynomial
	row := uint32(1)
	for n := 1; n < 32; n++ {
		odd[n] = row
		row <<= 1
	}
	gf2Square(&even, &odd)
	gf2Square(&odd, &even)
	for {
		gf2Square(&even, &odd)
		if len2&1 != 0 {
			crc1 = gf2Times(&even, crc1)
		}
		len2 >>= 1
		if len2 == 0 {
			break
		}
		gf2Square(&odd, &even)
		if len2&1 != 0 {
			crc1 = gf2Times(&odd, crc1)
		}
		len2 >>= 1
		if len2 == 0 {
			break
		}
	}
	return crc1 ^ crc2
}

func gf2Times(mat *[32]uint32, vec uint32) uint32 {
	var sum uint32
	for i := 0; vec != 0; i, vec = i+1, vec>>1 {
		if vec&1 != 0 {
			sum ^= mat[i]
		}
	}
	return sum
}

func gf2Square(square, mat *[32]uint32) {
	for n := 0; n < 32; n++ {
		square[n] = gf2Times(mat, mat[n])
	}
}
