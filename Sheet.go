package excel

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// Structs

// Sheet wraps the sheets of a excel file into a struct
type Sheet struct {
	file    *excelize.File
	name    string
	columns []string

	//experimenting
	draft     [][]Cell
	draftMode bool
}

// Get/Create Sheets

// Sheet retruns the sheet with the given name
func (excel *Excel) Sheet(name string) *Sheet {
	// Sheet exists
	for _, existingSheet := range *excel.sheets {
		if existingSheet.name == name {
			return &existingSheet
		}
	}

	newSheet := Sheet{file: excel.file, name: name, columns: []string{}, draft: [][]Cell{}, draftMode: false}
	excel.file.NewSheet(name)
	return &newSheet
}

// NewDraftSheet returns a sheet, that has draftmode enabled
func (excel *Excel) NewDraftSheet(name string) *Sheet {
	return &Sheet{file: excel.file, name: name, columns: []string{}, draft: [][]Cell{}, draftMode: true}
}

// FirstSheet returns the first sheet found in the excel file
func (excel *Excel) FirstSheet() *Sheet {
	shs := excel.sheets
	return &(*shs)[0]
}

// Filter Sheets

// FilterByHeader filters the excel file by its headertitle
func (sh *Sheet) FilterByHeader(header []string) [][]string {
	if sh.isEmpty() {
		return nil
	}

	data := sh.file.GetRows(sh.name)
	m := map[string]int{}

	for i, col := range data[0] {
		if contains(header, col) {
			m[col] = i
		}
	}
	sortedColumns := []string{}
	for _, h := range header {
		sortedColumns = append(sortedColumns, excelize.ToAlphaString(m[h]))
	}
	return sh.FilterByColumn(sortedColumns)
}

// FilterByColumn filters the excel file by its column
func (sh *Sheet) FilterByColumn(columns []string) [][]string {
	if sh.isEmpty() {
		return nil
	}
	data := sh.file.GetRows(sh.name)
	filteredData := [][]string{}

	for _, row := range data {
		filterMap := map[string]string{}
		for col, val := range row {
			if contains(columns, excelize.ToAlphaString(col)) {
				filterMap[excelize.ToAlphaString(col)] = val
			}
		}
		sortedRow := []string{}
		for _, c := range columns {
			sortedRow = append(sortedRow, filterMap[c])
		}
		filteredData = append(filteredData, sortedRow)
	}

	return filteredData[1:]
}

// Modify Sheets

// NextRow returns the next free Row
func (sh *Sheet) NextRow() int {
	if sh.draftMode {
		return len(sh.draft)
	}
	return sh.CurrentRow() + 1
}

// CurrentRow returns the current Row
func (sh *Sheet) CurrentRow() int {
	if sh.draftMode {
		return len(sh.draft) - 1
	}
	return len(sh.file.GetRows(sh.name))
}

// AddValue adds a value to the provided coordinates
func (sh *Sheet) AddValue(coords Coordinates, value interface{}, style Style) {
	sh.file.SetCellValue(sh.name, coords.ToString(), value)
	styleString := style.toString()
	if styleString == "" {
		return
	}
	st, err := sh.file.NewStyle(styleString)
	if err != nil {
		fmt.Println(styleString)
		fmt.Println(err)
	}
	sh.file.SetCellStyle(sh.name, coords.ToString(), coords.ToString(), st)
}

// AddRow scanns for the next available row and inserts cells at the given indexes provided by the map
func (sh *Sheet) AddRow(columnCellMap map[int]Cell) {
	freeRow := sh.NextRow()
	if sh.draftMode {
		columns := []int{}
		for col := range columnCellMap {
			columns = append(columns, col)
		}
		newRow := []Cell{}
		for i := 0; i != maxInt(columns); i++ {
			if val, ok := columnCellMap[i]; ok {
				newRow = append(newRow, val)
			} else {
				newRow = append(newRow, Cell{Value: draftCell, Style: NoStyle()})
			}
		}
		sh.draft = append(sh.draft, newRow)
		return
	}
	for col, cell := range columnCellMap {
		coords := Coordinates{Column: col, Row: freeRow}
		sh.file.SetCellValue(sh.name, coords.ToString(), cell.Value)
		styleString := cell.Style.toString()
		if styleString == "" {
			continue
		}
		st, err := sh.file.NewStyle(styleString)
		if err != nil {
			fmt.Println(styleString)
			fmt.Println(err)
		}
		sh.file.SetCellStyle(sh.name, coords.ToString(), coords.ToString(), st)
	}
}

// AddEmptyRow adds an empty row at index row
func (sh *Sheet) AddEmptyRow() {
	if sh.draftMode {
		sh.draft = append(sh.draft, []Cell{Cell{Value: " ", Style: NoStyle()}})
		return
	}
	freeRow := sh.NextRow()
	sh.file.SetCellStr(sh.name, Coordinates{Column: 0, Row: freeRow}.ToString(), " ")
}

// AddCondition adds a condition, that fills the cell red if its value is less than comparison
func (sh *Sheet) AddCondition(coord Coordinates, comparison float32) {
	compString := fmt.Sprintf("%f", comparison)
	format, err := sh.file.NewConditionalStyle(`{"fill":{"type":"pattern","color":["#F44E42"],"pattern":1}}`)
	if err != nil {
		fmt.Printf("couldn't create conditional style: %s\n", err)
	}
	sh.file.SetConditionalFormat(sh.name, coord.ToString(), fmt.Sprintf(`[{"type":"cell","criteria":"<","format":%d,"value":%s}]`, format, compString))
}

// GetValue returns the Value from the cell at coord
func (sh *Sheet) GetValue(coord Coordinates) interface{} {
	if sh.draftMode {
		return sh.draft[coord.Column][coord.Row].Value
	}
	return sh.file.GetCellValue(sh.name, coord.ToString())
}

// FreezeHeader freezes the headerrow
func (sh *Sheet) FreezeHeader() {
	sh.file.SetPanes(sh.name, `{"freeze":true,"split":false,"x_split":0,"y_split":1,"top_left_cell":"A34","active_pane":"bottomLeft"}`)
}

// Helper

func (sh *Sheet) isEmpty() bool {
	if sh.draftMode {
		if len(sh.draft) == 0 {
			return true
		}
		return false
	}
	if len(sh.file.GetRows(sh.name)) == 0 {
		return true
	}
	return false
}

// PrintHeader prints a table that contains the header of each sheet and it's index
func PrintHeader(sh *Sheet, startingRow int) {
	if sh.isEmpty() {
		return
	}
	sheetMap := sh.file.GetSheetMap()
	for k, v := range sheetMap {
		headerTableData := [][]string{}
		headerTableData = append(headerTableData, []string{strconv.Itoa(k), v})
		rows := sh.file.GetRows(v)
		for index, head := range rows[startingRow] {
			headerTableData = append(headerTableData, []string{fmt.Sprintf("%s%d", excelize.ToAlphaString(index), startingRow+1), head})
		}
		t := Table(headerTableData)
		fmt.Print(t)
		fmt.Println()
	}
}

// HeaderColumns returns the header columns of sheet
func (sh *Sheet) HeaderColumns() []string {
	return sh.columns
}

func contains(slice []string, value string) bool {
	for _, v := range slice {
		if v == value {
			return true
		}
	}
	return false
}

func (sh *Sheet) header() []string {
	if sh.draftMode {
		cols := []string{}
		for _, header := range sh.draft[0] {
			cols = append(cols, fmt.Sprintf("%v", header.Value))
		}
		return cols
	}
	return sh.file.GetRows(sh.name)[0]
}

func maxInt(slice []int) int {
	if len(slice) == 0 {
		fmt.Println("slice is empty")
		return 0
	}
	max := 0
	for _, i := range slice {
		if i > max {
			max = i
		}
	}
	return max
}
