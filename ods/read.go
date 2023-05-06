// This package implements rudimentary support
// for reading Open Document Spreadsheet files. At current
// stage table data can be accessed.
package ods

import (
	"bytes"
	"encoding/xml"
	"errors"
	"image/color"
	"io"
	"log"
	"strconv"
	"strings"

	"github.com/knieriem/odf"
)

type Doc struct {
	XMLName xml.Name `xml:"document-content"`
	Table   []Table  `xml:"body>spreadsheet>table"`
	Style   []Style  `xml:"automatic-styles>style"`
}

type Style struct {
	Name           string     `xml:"name,attr"`
	ColumnProps    SCol       `xml:"table-column-properties"`
	RowProps       SRow       `xml:"table-row-properties"`
	CellProps      SCell      `xml:"table-cell-properties"`
	TextProps      SText      `xml:"text-properties"`
	ParagraphProps SParagraph `xml:"paragraph-properties"`
}

type SCol struct {
	Width       string `xml:"column-width,attr"`
	BreakBefore string `xml:"break-before,attr"`
	Cell
}

type SRow struct {
	Height        string `xml:"row-height,attr"`
	BreakBefore   string `xml:"break-before,attr"`
	OptimalHeight bool   `xml:"use-optimal-row-height,attr"`
}

type SCell struct {
	BorderTop       string `xml:"border-top,attr"`
	BorderBottom    string `xml:"border-bottom,attr"`
	BorderLeft      string `xml:"border-left,attr"`
	BorderRight     string `xml:"border-right,attr"`
	BackgroundColor string `xml:"background-color,attr"`
	AlignVertical   string `xml:"vertical-align,attr"`
	NrColsSpanned   int    `xml:"number-columns-spanned"`
	NrRowsSpanned   int    `xml:"number-rows-spanned"`
	Padding         string `xml:"padding"`
}

type SText struct {
	Name   string `xml:"font-name,attr"`
	Size   string `xml:"font-size,attr"`
	Weight string `xml:"font-weight,attr"`
	Color  string `xml:"color"`
}

type SParagraph struct {
	Align      string `xml:"text-align,attr"`
	MarginLeft string `xml:"margin-left,attr"`
}

type Table struct {
	Name      string    `xml:"name,attr"`
	XMLColumn []TColumn `xml:"table-column"`
	XMLRow    []TRow    `xml:"table-row"`
}

// Row type is processed row, contaning cells.
// This is struct user is working with.
type Row struct {
	Cell []Cell
}

// Cell type is processed cell
type Cell struct {
	FontName        string
	FontSize        float64
	FontWeight      string
	FontColor       color.RGBA
	BackgroundColor color.RGBA
	Value           string
	// Align can be: "start", "center", "end".
	Align string
	// AlignVertical can be: "top", "middle", "bottom".
	AlignVertical string
	// Width is in mm
	Width float64
	// Height is in mm
	Height float64
	// Padding in mm
	Padding float64
}

type TColumn struct {
	RepeatedCols    int    `xml:"number-columns-repeated,attr"`
	StyleName       string `xml:"style-name,attr"`
	DefaltCellStyle string `xml:"default-cell-style-name,attr"`
}

type TRow struct {
	RepeatedRows int    `xml:"number-rows-repeated,attr"`
	StyleName    string `xml:"style-name,attr"`

	Cell []TCell `xml:",any"` // use ",any" to match table-cell and covered-table-cell
}

func (r *TRow) IsEmpty() bool {
	for _, c := range r.Cell {
		if !c.IsEmpty() {
			return false
		}
	}
	return true
}

// WidthInMM skips cells from the beginning
// which are empty
func (r *Row) WidthInMM() float64 {
	sum := 0.0
	prevEmpty := true
	for i := range r.Cell {
		empty := r.Cell[i].IsEmpty()
		if !(prevEmpty && empty) {
			sum += r.Cell[i].Width
		}
		if empty { // if ever non-empty - lock it
			prevEmpty = empty
		}
	}
	return sum
}

func (r *Row) HeightInMM() float64 {
	if len(r.Cell) > 0 {
		return r.Cell[0].Height
	}
	return -1
}

// Return the contents of a row as a slice of strings. Cells that are
// covered by other cells will appear as empty strings.
func (r *TRow) Strings(b *bytes.Buffer) (row []string) {
	n := len(r.Cell)
	if n == 0 {
		return
	}

	// remove trailing empty cells
	for i := n - 1; i >= 0; i-- {
		if !r.Cell[i].IsEmpty() {
			break
		}
		n--
	}
	r.Cell = r.Cell[:n]

	n = 0
	// calculate the real number of cells (including repeated)
	for _, c := range r.Cell {
		switch {
		case c.RepeatedCols != 0:
			n += c.RepeatedCols
		default:
			n++
		}
	}

	row = make([]string, n)
	w := 0
	for _, c := range r.Cell {
		cs := ""
		if c.XMLName.Local != "covered-table-cell" {
			cs = c.PlainText(b)
		}
		row[w] = cs
		w++
		switch {
		case c.RepeatedCols != 0:
			for j := 1; j < c.RepeatedCols; j++ {
				row[w] = cs
				w++
			}
		}
	}
	return
}

func (r *TRow) Cells(b *bytes.Buffer, styles []Style, sRow SRow, tColumns []TColumn) Row {
	n := len(r.Cell)
	if n == 0 {
		return Row{}
	}

	// remove trailing empty cells
	for i := n - 1; i >= 0; i-- {
		if !r.Cell[i].IsEmpty() {
			break
		}
		n--
	}
	r.Cell = r.Cell[:n]

	n = 0
	// calculate the real number of cells (including repeated)
	for _, c := range r.Cell {
		switch {
		case c.RepeatedCols != 0:
			n += c.RepeatedCols
		default:
			n++
		}
	}

	//DEBUG
	//log.Println("cols len:", len(tColumns))
	//for i := range tColumns {
	//	log.Println("repeated:", tColumns[i])
	//}
	//for i := range styles {
	//	log.Println("style, col", i, " width:", styles[i].ColumnProps.Width)
	//}

	cels := make([]Cell, n)
	w := 0
	nr := 0
	rs := 0
	for _, c := range r.Cell { // i = possition in the row
		plain := ""
		if c.XMLName.Local != "covered-table-cell" {
			plain = c.PlainText(b)
		}

		// fix cell width if spanned
		// it must be calculated before incrementing nr var
		// PROBLEM: cols may be repeated, so nr+1 may not work
		sum := 0.0
		if c.ColSpan != 0 {
			for i := 0; i < c.ColSpan; i++ {
				wid := 0.0
				var err error
				if i < tColumns[nr].RepeatedCols {
					wid, err = ToMM(GetColStyleByName(tColumns[nr].StyleName, styles).Width)
				} else {
					wid, err = ToMM(GetColStyleByName(tColumns[nr+i].StyleName, styles).Width)
				}
				if err != nil {
					log.Println("TRow.Cells:", err)
				}
				sum += wid
			}
		}

		// here we are deciding which column style
		// we should use for current cell
		// columns can also repeat
		coln := tColumns[nr]
		if tColumns[nr].RepeatedCols == 0 {
			nr++
			rs = 0
		} else {
			rs++
			maxrs := tColumns[nr].RepeatedCols
			if maxrs == rs {
				nr++
			}
		}
		sCol := GetColStyleByName(coln.StyleName, styles)
		sCell := GetCellStyleByName(c.StyleName, styles)
		sDefaultColCell := GetCellStyleByName(coln.DefaltCellStyle, styles)
		cell := ConsolidateStyles(sRow, sCol, sCell, sDefaultColCell)
		cell.Value = plain

		if c.ColSpan != 0 {
			cell.Width = sum
		}

		cels[w] = cell
		w++
		if c.RepeatedCols != 0 {
			for j := 1; j < c.RepeatedCols; j++ {
				cels[w] = cell
				w++
			}
		}
	}
	return Row{Cell: cels}
}

type TCell struct {
	XMLName xml.Name

	// attributes
	ValueType    string `xml:"value-type,attr"`
	Value        string `xml:"value,attr"`
	Formula      string `xml:"formula,attr"`
	RepeatedCols int    `xml:"number-columns-repeated,attr"`
	ColSpan      int    `xml:"number-columns-spanned,attr"`
	RowSpan      int    `xml:"number-rows-spanned,attr"`
	StyleName    string `xml:"style-name,attr"`

	P []Par `xml:"p"`
}

func (c *TCell) IsEmpty() (empty bool) {
	switch len(c.P) {
	case 0:
		empty = true
	case 1:
		if c.P[0].XML == "" {
			empty = true
		}
	}
	return
}

// PlainText extracts the text from a cell. Space tags (<text:s text:c="#">)
// are recognized. Inline elements (like span) are ignored, but the
// text they contain is preserved
func (c *TCell) PlainText(b *bytes.Buffer) string {
	n := len(c.P)
	if n == 1 {
		return c.P[0].PlainText(b)
	}

	b.Reset()
	for i := range c.P {
		if i != n-1 {
			c.P[i].writePlainText(b)
			b.WriteByte('\n')
		} else {
			c.P[i].writePlainText(b)
		}
	}
	return b.String()
}

type Par struct {
	XML string `xml:",innerxml"`
}

func (p *Par) PlainText(b *bytes.Buffer) string {
	for i := range p.XML {
		if p.XML[i] == '<' || p.XML[i] == '&' {
			b.Reset()
			p.writePlainText(b)
			return b.String()
		}
	}
	return p.XML
}
func (p *Par) writePlainText(b *bytes.Buffer) {
	for i := range p.XML {
		if p.XML[i] == '<' || p.XML[i] == '&' {
			goto decode
		}
	}
	b.WriteString(p.XML)
	return

decode:
	d := xml.NewDecoder(strings.NewReader(p.XML))
	for {
		t, _ := d.Token()
		if t == nil {
			break
		}
		switch el := t.(type) {
		case xml.StartElement:
			switch el.Name.Local {
			case "s":
				n := 1
				for _, a := range el.Attr {
					if a.Name.Local == "c" {
						n, _ = strconv.Atoi(a.Value)
					}
				}
				for i := 0; i < n; i++ {
					b.WriteByte(' ')
				}
			}
		case xml.CharData:
			b.Write(el)
		}
	}
}

func (t *Table) Width() int {
	return len(t.XMLColumn)
}
func (t *Table) Height() int {
	return len(t.XMLRow)
}

// getRowsNr removes trailing empty rows,
// returns real number of rows (includes empty rows before table)
func (t *Table) removeTrailingEmptyRows() int {

	n := len(t.XMLRow)
	if n == 0 {
		return n
	}

	// remove trailing empty rows
	for i := n - 1; i >= 0; i-- {
		if !t.XMLRow[i].IsEmpty() {
			break
		}
		n--
	}
	t.XMLRow = t.XMLRow[:n]

	n = 0
	// calculate the real number of rows (including repeated rows)
	for _, r := range t.XMLRow {
		switch {
		case r.RepeatedRows != 0:
			n += r.RepeatedRows
		default:
			n++
		}
	}

	return n
}

func (t *Table) Strings() (s [][]string) {
	var b bytes.Buffer

	n := t.removeTrailingEmptyRows()

	s = make([][]string, n)
	w := 0
	for _, r := range t.XMLRow {
		row := r.Strings(&b)
		s[w] = row
		w++
		for j := 1; j < r.RepeatedRows; j++ {
			s[w] = row
			w++
		}
	}
	return
}

func (c *Cell) IsEmpty() bool {
	if c.Value == "" {
		return true
	}
	return false
}

func (r *Row) IsEmpty() bool {
	for _, c := range r.Cell {
		if !c.IsEmpty() {
			return false
		}
	}
	return true
}

func (t *Table) Rows(styles []Style) (rr []Row) {
	var b bytes.Buffer

	n := t.removeTrailingEmptyRows()

	rr = make([]Row, n)
	w := 0
	for _, r := range t.XMLRow {
		row := r.Cells(&b, styles, GetRowStyleByName(r.StyleName, styles), t.XMLColumn)
		rr[w] = row
		w++
		for j := 1; j < r.RepeatedRows; j++ {
			rr[w] = row
			w++
		}
	}
	return
}

func GetRowStyleByName(name string, styles []Style) SRow {
	for i := range styles {
		if styles[i].Name == name {
			return styles[i].RowProps
		}
	}
	return SRow{}
}

func GetColStyleByName(name string, styles []Style) SCol {
	for i := range styles {
		if styles[i].Name == name {
			return styles[i].ColumnProps
		}
	}
	return SCol{}
}

func GetCellStyleByName(name string, styles []Style) Style {
	for i := range styles {
		if styles[i].Name == name {
			return styles[i]
		}
	}
	return Style{}
}

// ConsolidateStyles - TODO: add all params
func ConsolidateStyles(r SRow, c SCol, cell, defaultColCell Style) Cell {
	var err error

	w, err := ToMM(c.Width)
	if err != nil {
		log.Println(err)
	}
	h, err := ToMM(r.Height)
	if err != nil {
		log.Println(err)
	}

	bgColor := color.RGBA{}
	if cell.CellProps.BackgroundColor != "" {
		bgColor, err = ParseHexColor(cell.CellProps.BackgroundColor)
	} else {
		bgColor, err = ParseHexColor(defaultColCell.CellProps.BackgroundColor)
	}
	if err != nil {
		log.Println("error parsing background hex to RGBA,", err)
	}

	fontColor := color.RGBA{}
	if cell.TextProps.Color != "" {
		fontColor, err = ParseHexColor(cell.TextProps.Color)
	} else {
		fontColor, err = ParseHexColor(defaultColCell.TextProps.Color)
	}
	if err != nil {
		log.Println("error parsing text color hex to RGBA,", err)
	}

	fontSize := 0.0
	if cell.TextProps.Size != "" {
		fontSize, err = PxToFloat64(cell.TextProps.Size)
	} else {
		fontSize, err = PxToFloat64(defaultColCell.TextProps.Size)
	}
	if err != nil {
		log.Println(err)
	}

	align := ""
	if cell.ParagraphProps.Align != "" {
		align = cell.ParagraphProps.Align
	} else {
		align = defaultColCell.ParagraphProps.Align
	}

	alignVert := ""
	if cell.CellProps.AlignVertical != "" {
		alignVert = cell.CellProps.AlignVertical
	} else {
		alignVert = defaultColCell.CellProps.AlignVertical
	}

	fontName := ""
	fontWeight := ""
	if cell.TextProps.Name != "" {
		fontName = cell.TextProps.Name
		fontWeight = cell.TextProps.Weight
	} else {
		fontName = defaultColCell.TextProps.Name
		fontWeight = defaultColCell.TextProps.Weight
	}

	padding := 0.0
	if cell.CellProps.Padding != "" {
		padding, err = ToMM(cell.CellProps.Padding)
	} else {
		padding, err = ToMM(defaultColCell.CellProps.Padding)
	}
	if err != nil {
		log.Println(err)
	}

	return Cell{
		Width:           w,
		Height:          h,
		Align:           align,
		AlignVertical:   alignVert,
		FontName:        fontName,
		FontSize:        fontSize,
		FontWeight:      fontWeight,
		FontColor:       fontColor,
		BackgroundColor: bgColor,
		Padding:         padding,
	}
}

type File struct {
	*odf.File
}

// Open an ODS file. If the file doesn't exist or doesn't look
// like a spreadsheet file, an error is returned.
func Open(fileName string) (*File, error) {
	f, err := odf.Open(fileName)
	if err != nil {
		return nil, err
	}
	return newFile(f)
}

// NewReader initializes a File struct with an already opened
// ODS file, and checks the spreadsheet's media type.
func NewReader(r io.ReaderAt, size int64) (*File, error) {
	f, err := odf.NewReader(r, size)
	if err != nil {
		return nil, err
	}
	return newFile(f)
}

func newFile(f *odf.File) (*File, error) {
	if f.MimeType != odf.MimeTypePfx+"spreadsheet" {
		f.Close()
		return nil, errors.New("not a spreadsheet")
	}
	return &File{f}, nil
}

// Parse the content.xml part of an ODS file. On Success
// the returned Doc will contain the data of the rows and cells
// of the table(s) contained in the ODS file.
func (f *File) ParseContent(doc *Doc) (err error) {
	content, err := f.Open("content.xml")
	if err != nil {
		return
	}
	defer content.Close()

	d := xml.NewDecoder(content)
	err = d.Decode(doc)
	return
}
