package excel

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// Structs

// Sheet wraps the sheets of a excel file into a struct
type Sheet struct {
	file        *excelize.File
	name        string
	columns     []string
	draft       [][]Cell
	writeAccess bool
}

// Get/Create Sheets

// Sheet retruns the sheet with the given name or creates a new one
func (excel *Excel) Sheet(name string) *Sheet {
	// Sheet exists
	for _, existingSheet := range *excel.sheets {
		if existingSheet.name == name {
			return &existingSheet
		}
	}

	newSheet := Sheet{file: excel.file, name: name, columns: []string{}, draft: [][]Cell{}, writeAccess: true}
	excel.file.NewSheet(name)
	return &newSheet
}

// ClearSheet clears the content of sheet, but not the draft of sheet
func (sh *Sheet) clearSheet() {
	name := sh.name
	sh.file.DeleteSheet(name)
	sh.file.NewSheet(name)
}

// GetWriteAccess populates draft with current content fo sheet and grants write access
func (sh *Sheet) GetWriteAccess() {
	if sh.writeAccess {
		fmt.Printf("write access for sheet %s already granted\n", sh.name)
		return
	}
	rows := sh.file.GetRows(sh.name)
	for i, row := range rows {
		newCellRow := []Cell{}
		for j, str := range row {
			styleID := sh.file.GetCellStyle(sh.name, Coordinates{Row: i + 1, Column: j + 1}.ToString())
			newCellRow = append(newCellRow, Cell{Value: str, Style: RawID(styleID)})
		}
		sh.draft = append(sh.draft, newCellRow)
	}
}

// FirstSheet returns the first sheet found in the excel file
func (excel *Excel) FirstSheet() *Sheet {
	shs := excel.sheets
	return &(*shs)[0]
}

// ExtractColumnsByName extracts columns by there names from sheet
func (sh *Sheet) ExtractColumnsByName(columnNames []string) [][]string {
	columns := []string{}
	columnMap := map[string]int{}
	for i, name := range sh.columns {
		if containsString(columnNames, name) {
			columnMap[name] = i + 1
		}
	}
	for _, columnName := range columnNames {
		columnstring, err := excelize.ColumnNumberToName(columnMap[columnName])
		if err != nil {
			fmt.Printf("error converting index to columnname: %s\n", err)
			continue
		}
		columns = append(columns, columnstring)
	}
	return sh.ExtractColumns(columns)
}

// ExtractColumns returns columns from sheet
func (sh *Sheet) ExtractColumns(columns []string) [][]string {
	numeric := []int{}
	rawData := sh.file.GetRows(sh.name)
	filteredData := [][]string{}

	for _, c := range columns {
		numeric = append(numeric, excelize.MustColumnNameToNumber(c))
	}
	for _, row := range rawData {
		filteredRow := []string{}
		for j, cell := range row {
			if containsInt(numeric, j+1) {
				filteredRow = append(filteredRow, cell)
			}
		}
		filteredData = append(filteredData, filteredRow)
	}
	return filteredData[1:]
}

// Modify Sheets

// NextRow returns the next free Row
func (sh *Sheet) NextRow() int {
	return sh.CurrentRow() + 1
}

// CurrentRow returns the current Row
func (sh *Sheet) CurrentRow() int {
	if !sh.writeAccess {
		return len(sh.file.GetRows(sh.name))
	}
	return len(sh.draft)
}

// AddHeaderColumn adds a header column to sheet
func (sh *Sheet) AddHeaderColumn(header []string) {
	if !sh.writeAccess {
		fmt.Printf("no permission to write to sheet %s\n", sh.name)
		return
	}

	for _, h := range header {
		sh.draft[0] = append(sh.draft[0], Cell{Value: h, Style: NoStyle()})
	}
	sh.columns = header
}

// AddRow scanns for the next available row and inserts cells at the given indexes provided by the map
func (sh *Sheet) AddRow(columnCellMap map[int]Cell) {
	if !sh.writeAccess {
		fmt.Printf("no permission to write to sheet %s\n", sh.name)
		return
	}

	if len(sh.draft) == 0 {
		fmt.Printf("WARNING: Sheet %s has no header column\n", sh.name)
	}
	columns := []int{}
	for col := range columnCellMap {
		columns = append(columns, col)
	}
	newRow := []Cell{}
	for i := 0; i < len(sh.columns); i++ {
		if val, ok := columnCellMap[i]; ok {
			newRow = append(newRow, val)
		} else {
			newRow = append(newRow, Cell{Value: draftCell, Style: NoStyle()})
		}
	}

	sh.draft = append(sh.draft, newRow)
}

// AddEmptyRow adds an empty row at index row
func (sh *Sheet) AddEmptyRow() {
	if !sh.writeAccess {
		fmt.Printf("no permission to write to sheet %s\n", sh.name)
		return
	}
	sh.draft = append(sh.draft, []Cell{Cell{Value: " ", Style: NoStyle()}})
}

// AddCondition adds a condition, that fills the cell red if its value is less than comparison
// func (sh *Sheet) AddCondition(coord Coordinates, comparison float32) {
// 	compString := fmt.Sprintf("%f", comparison)
// 	format, err := sh.file.NewConditionalStyle(`{"fill":{"type":"pattern","color":["#F44E42"],"pattern":1}}`)
// 	if err != nil {
// 		fmt.Printf("couldn't create conditional style: %s\n", err)
// 	}
// 	sh.file.SetConditionalFormat(sh.name, coord.ToString(), fmt.Sprintf(`[{"type":"cell","criteria":"<","format":%d,"value":%s}]`, format, compString))
// }

// GetValue returns the Value from the cell at coord
func (sh *Sheet) GetValue(coord Coordinates) interface{} {
	if !sh.writeAccess {
		return sh.file.GetCellValue(sh.name, coord.ToString())
	}
	return sh.draft[coord.Column][coord.Row].Value
}

// FreezeHeader freezes the headerrow
func (sh *Sheet) FreezeHeader() {
	sh.file.SetPanes(sh.name, `{"freeze":true,"split":false,"x_split":0,"y_split":1,"top_left_cell":"A34","active_pane":"bottomLeft"}`)
}

// Helper

func (sh *Sheet) isEmpty() bool {
	if len(sh.draft) == 0 {
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
			coordString, err := excelize.CoordinatesToCellName(index, startingRow+1)
			if err != nil {
				fmt.Println(err)
			}
			headerTableData = append(headerTableData, []string{coordString, head})
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

func containsInt(slice []int, value int) bool {
	for _, i := range slice {
		if i == value {
			return true
		}
	}
	return false
}

func containsString(slice []string, value string) bool {
	for _, v := range slice {
		if v == value {
			return true
		}
	}
	return false
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
