package xls

import (
	"io"
	"os"

	"github.com/extrame/ole2"
)

//Open one xls file
// The returned WorkBook now supports Close() to release resources.
// IMPORTANT: Call wb.Close() when done to prevent memory leaks.
func Open(file string, charset string) (*WorkBook, error) {
	fi, err := os.Open(file)
	if err != nil {
		return nil, err
	}

	wb, err := openReaderWithCloser(fi, charset, fi)
	if err != nil {
		fi.Close() // Close on error
		return nil, err
	}
	return wb, nil
}

//Open one xls file and return the closer
// Deprecated: Use Open() instead, which now has Close() method on WorkBook.
// This function is kept for backwards compatibility.
func OpenWithCloser(file string, charset string) (*WorkBook, io.Closer, error) {
	fi, err := os.Open(file)
	if err != nil {
		return nil, nil, err
	}

	wb, err := openReaderWithCloser(fi, charset, fi)
	if err != nil {
		fi.Close()
		return nil, nil, err
	}
	return wb, fi, nil
}

//Open xls file from reader
// Note: When using OpenReader directly, you are responsible for closing
// any underlying file handles. The WorkBook will NOT automatically close them.
// Use Open() instead for automatic resource management.
func OpenReader(reader io.ReadSeeker, charset string) (wb *WorkBook, err error) {
	return openReaderWithCloser(reader, charset, nil)
}

// Internal function that opens a reader and optionally stores a closer
func openReaderWithCloser(reader io.ReadSeeker, charset string, closer io.Closer) (wb *WorkBook, err error) {
	var ole *ole2.Ole
	if ole, err = ole2.Open(reader, charset); err == nil {
		var dir []*ole2.File
		if dir, err = ole.ListDir(); err == nil {
			var book *ole2.File
			var root *ole2.File
			for _, file := range dir {
				name := file.Name()
				if name == "Workbook" {
					if book == nil {
						book = file
					}
					//book = file
					// break
				}
				if name == "Book" {
					book = file
					// break
				}
				if name == "Root Entry" {
					root = file
				}
			}
			if book != nil {
				wb = newWorkBookFromOle2(ole.OpenFile(book, root), closer)
				return
			}
		}
	}
	return
}
