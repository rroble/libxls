package main

// #cgo LDFLAGS: -L.libs -lxlsreader
// #include <include/xls.h>
// #include <stdlib.h>
import "C"
import (
	"fmt"
	"io/ioutil"
	"unsafe"
)

type summaryInfo struct {
	title      string
	subject    string
	author     string
	keywords   string
	comment    string
	lastAuthor string
	appname    string
	category   string
	manager    string
	company    string
}

func main() {
	fmt.Println("~START~")
	defer fmt.Println("~END-")
	dat, e := ioutil.ReadFile("test/files/test2.xls")
	if e != nil {
		panic(e)
	}
	data := C.CBytes(dat)
	charset := C.CString("utf-8")

	defer func() {
		C.free(unsafe.Pointer(data))
		C.free(unsafe.Pointer(charset))
	}()

	var err C.xls_error_t
	wb := C.xls_open_buffer((*C.uchar)(data), (C.ulong)(len(dat)), (*C.char)(charset), &err)
	if err != C.LIBXLS_OK {
		panic(err)
	}
	defer C.xls_close(wb)
	// wb
	// .sheets.count
	// .sheets.sheet[i].name

	info := C.xls_summaryInfo(wb)
	defer C.xls_close_summaryInfo(info)
	// fmt.Printf("%#v\n", summary(info))

	// fmt.Println("Sheets", wb.sheets.count)

	for i := 0; ; i++ {
		sheet := C.xls_getWorkSheet(wb, C.int(i))
		if sheet == nil {
			break
		}
		err := C.xls_parseWorkSheet(sheet)
		if err != C.LIBXLS_OK {
			panic(err)
		}
		fmt.Println("Sheet ", i /*, wb.sheets.sheet[i].name*/)
		lastRow := int(sheet.rows.lastrow)
		for j := 0; j <= lastRow; j++ {
			cellRow := (C.WORD)(j)
			fmt.Print("Row ", cellRow+1, ": ")

			var cellCol C.WORD
			for cellCol = 0; cellCol <= sheet.rows.lastcol; cellCol++ {
				cell := C.xls_cell(sheet, cellRow, cellCol)
				if cell == nil || cell.isHidden == 1 {
					continue
				}
				if cell.rowspan > 1 {
					// set next row/s cell value
				}

				if cell.id == C.XLS_RECORD_RK || cell.id == C.XLS_RECORD_MULRK || cell.id == C.XLS_RECORD_NUMBER {
					fmt.Print("", cell.d, ", ")
				} else if cell.id == C.XLS_RECORD_FORMULA || cell.id == C.XLS_RECORD_FORMULA_ALT {
					if cell.l == 0 {
						fmt.Print("", cell.d, ", ")
					} else {
						str := C.GoString(cell.str)
						if str == "bool" {
							if cell.d == 0 {
								fmt.Print("false, ")
							} else {
								fmt.Print("true, ")
							}
						} else if str == "error" {
							fmt.Print("*error*", cell.d, ", ")
						} else {
							fmt.Print("", C.GoString(cell.str), ", ")
						}
					}
				} else if cell.str != nil {
					fmt.Print("", C.GoString(cell.str), ", ")
				} else {
					fmt.Print(", ")
				}
			}
			fmt.Println("")
		}
		C.xls_close_WS(sheet)
	}
}

func summary(info *C.xlsSummaryInfo) summaryInfo {
	return summaryInfo{
		C.GoString((*C.char)(unsafe.Pointer(info.title))),
		C.GoString((*C.char)(unsafe.Pointer(info.subject))),
		C.GoString((*C.char)(unsafe.Pointer(info.author))),
		C.GoString((*C.char)(unsafe.Pointer(info.keywords))),
		C.GoString((*C.char)(unsafe.Pointer(info.comment))),
		C.GoString((*C.char)(unsafe.Pointer(info.lastAuthor))),
		C.GoString((*C.char)(unsafe.Pointer(info.appName))),
		C.GoString((*C.char)(unsafe.Pointer(info.category))),
		C.GoString((*C.char)(unsafe.Pointer(info.manager))),
		C.GoString((*C.char)(unsafe.Pointer(info.company))),
	}
}

// macro? autoconf-archive
// ./configure --disable-shared --enable-static
