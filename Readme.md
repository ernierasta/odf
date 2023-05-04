This fork is an attempt to get (almost) full information
from ods file.

You can now get height, width of the cell, determine
style (font, alignment, colors, ...).

I will probably not merge it to upstream as my plan is
to rewrite it, so it will not make too much sense for
original author to maintain it.

Quickstart (error handling ommited!):
```go
var doc Doc

f, _ := Open("./table1.ods")
defer f.Close()
_ := f.ParseContent(&doc)

rows := doc.Table[0].Rows(doc.Style)
for i := range rr {
    fmt.Println("row", i)
    for n := range rr[i].Cell {
        fmt.Println("value:", rr[i].Cell[n].Value)
        fmt.Println("width:", rr[i].Cell[n].Width, "height:", rr[i].Cell[n].Height)
    }
}
```


ORIGINAL DESCRIPTION:

This projekt contains two Go packages – odf and odf/ods
– that allow basic read-only access to the tables of Open
Document Spreadsheets, making use of Go's encoding/xml package.

For now the ods package makes it easy to convert a table to a
`[][]string`.
