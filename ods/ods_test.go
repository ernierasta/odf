package ods

import (
	"fmt"
	"os"
	"strconv"
	"testing"
)

// "A"	"1"	"A cell containing\nmore than one line."
// "B"	"foo"	""
// ""	"4"	"quote\"quote"
// "14.01.12"	""	""
// ""	""	""
// "cell spanning two columns"	""	""
// ""	""	""
// ""	""	""
// "aaa"	"cell spanning two rows"	"ccc"
// "aa"	""	"cc"
// ""	""	""
// "same content"	"same content"	"same content"
// ""	""	""
// "same content"	""	""
// "same content"	""	""
// "same content"	""	""
// ""	""	""
// "Cell with  styles"	""	""
func ExampleParsing() {
	var doc Doc

	f, err := Open("./test.ods")
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		return
	}
	defer f.Close()
	if err := f.ParseContent(&doc); err != nil {
		fmt.Fprintln(os.Stderr, err)
		return
	}

	// Dump the first table one line per row, writing
	// tab separated, quoted fields.
	if len(doc.Table) > 0 {
		for _, row := range doc.Table[0].Strings() {
			sep := ""
			for _, field := range row {
				fmt.Print(sep, strconv.Quote(field))
				sep = "\t"
			}
			fmt.Print("\n")
		}
	}
}

func TestDummy(_ *testing.T) {
}
