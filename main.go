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

	info := C.xls_summaryInfo(wb)
	defer C.xls_close_summaryInfo(info)
	fmt.Printf("%#v\n", summary(info))
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

// LD_LIBRARY_PATH=.libs:$LD_LIBRARY_PATH go run main.go
