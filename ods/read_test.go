package ods

import (
	"fmt"
	"os"
	"testing"

	"github.com/davecgh/go-spew/spew"
)

func TestParse(t *testing.T) {
	var doc Doc

	f, err := Open("./test2.ods")
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		return
	}
	defer f.Close()
	if err := f.ParseContent(&doc); err != nil {
		fmt.Fprintln(os.Stderr, err)
		return
	}

	// dump all
	//spew.Dump(doc.Table[0].XMLRow)
	//fmt.Println("cols:")
	//spew.Dump(doc.Table[0].XMLColumn)

	rr := doc.Table[0].Rows(doc.Style)
	//spew.Dump(rr)
	for i := range rr {
		fmt.Println("row", i)
		for n := range rr[i].Cell {
			fmt.Println("value:", rr[i].Cell[n].Value)
			//spew.Dump(rr[i].Cell[n])
			fmt.Println("width:", rr[i].Cell[n].Width, "height:", rr[i].Cell[n].Height)
			spew.Dump(rr[i].Cell[n])
		}
	}
}
