package excel

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// Structs

// Sheet wraps the sheets of a excel file into a struct
type Sheet struct {
	file         *excelize.File
	name         string
	columns      []string
	draft        [][]Cell
	writeAccess  bool
	freezeHeader bool
}

// Get/Create Sheets

// Sheet retruns the sheet with the given name or creates a new one
func (excel *Excel) Sheet(name string) *Sheet {
	// Sheet exists
	for i, existingSheet := range *excel.sheets {
		if existingSheet.name == name {
			return &(*excel.sheets)[i]
		}
	}
	fmt.Printf("Creating new sheet %s\n", name)
	newSheet := Sheet{file: excel.file, name: name, columns: []string{}, draft: [][]Cell{}, writeAccess: true}
	excel.file.NewSheet(name)
	*excel.sheets = append(*excel.sheets, newSheet)
	return &(*excel.sheets)[len(*excel.sheets)-1]
}

// ClearSheet clears the content of sheet, but not the draft of sheet
func (sh *Sheet) clearSheet() {
	name := sh.name
	sh.file.DeleteSheet(name)
	sh.file.NewSheet(name)
}

// Name returns the name of sheet
func (sh *Sheet) Name() string {
	return sh.name
}

// Draft returns a copy of the current draft of sheet
func (sh *Sheet) Draft() [][]Cell {
	return sh.draft
}

// GetWriteAccess populates draft with current content fo sheet and grants write access
func (sh *Sheet) GetWriteAccess() {
	if sh.writeAccess {
		fmt.Printf("write access for sheet %s already granted\n", sh.name)
		return
	}
	sh.draft = [][]Cell{}
	rows, _ := sh.file.GetRows(sh.name)
	for i, row := range rows {
		newCellRow := []Cell{}
		for j, str := range row {
			styleID, _ := sh.file.GetCellStyle(sh.name, Coordinates{Row: i + 1, Column: j + 1}.String())
			newCellRow = append(newCellRow, Cell{Value: str, Style: RawID(styleID)})
		}
		sh.draft = append(sh.draft, newCellRow)
	}
	for _, cell := range sh.draft[0] {
		sh.columns = append(sh.columns, fmt.Sprintf("%s", cell.Value))
	}
	sh.writeAccess = true
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
	rawData, _ := sh.file.GetRows(sh.name)
	filteredData := [][]string{}

	for _, c := range columns {
		num, _ := excelize.ColumnNameToNumber(c)
		numeric = append(numeric, num)
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
		rows, err := sh.file.GetRows(sh.name)
		if err != nil {
			fmt.Println(err)
		}
		return len(rows)
	}
	return len(sh.draft)
}

// AddHeaderColumn adds a header column to sheet
func (sh *Sheet) AddHeaderColumn(header []string) {
	if !sh.writeAccess {
		fmt.Printf("no permission to write to sheet %s\n", sh.name)
		return
	}

	headerCells := []Cell{}
	for i, h := range header {
		headerCells = append(headerCells, Cell{Value: h, Style: NoStyle(), coordinates: Coordinates{Column: i + 1, Row: 1}})
	}
	if len(sh.draft) == 0 {
		fmt.Println("Writing Header Column:")
		sh.draft = append(sh.draft, headerCells)
	} else {
		fmt.Println("Replacing Header Column:")
		sh.draft[0] = headerCells
	}
	sh.columns = header
	fmt.Println(sh.columns)
}

// AddRow scanns for the next available row and inserts cells at the given indexes provided by the map
func (sh *Sheet) AddRow(columnCellMap map[int]Cell) {
	if !sh.writeAccess {
		fmt.Printf("no permission to write to sheet %s\n", sh.name)
		return
	}

	newRowIndexes := []int{}
	for index := range columnCellMap {
		newRowIndexes = append(newRowIndexes, index)
	}
	newRow := []Cell{}

	for i := 1; i != maxInt(newRowIndexes)+1; i++ {
		if val, ok := columnCellMap[i]; ok {
			val.coordinates = Coordinates{Column: i, Row: len(sh.draft)}
			str := strings.TrimSpace(fmt.Sprintf("%s", val.Value))
			if str == "" {
				val.Value = StyleCell
			}
			newRow = append(newRow, val)
		} else {
			newRow = append(newRow, Cell{Value: DraftCell, Style: NoStyle(), coordinates: Coordinates{Column: i, Row: len(sh.draft)}})
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
	sh.draft = append(sh.draft, []Cell{Cell{Value: DraftCell, Style: NoStyle(), coordinates: Coordinates{Column: 1, Row: len(sh.draft)}}})
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

// CopyRow appends row from sheet to the draft of the calling sheet
func (sh *Sheet) CopyRow(sheet *Sheet, row int) {
	sh.draft = append(sh.draft, sheet.draft[row])
}

// GetValue returns the Value from the cell at coord
func (sh *Sheet) GetValue(coord Coordinates) interface{} {
	if !sh.writeAccess {
		value, err := sh.file.GetCellValue(sh.name, coord.String())
		if err != nil {
			fmt.Println(err)
		}
		return value
	}
	fmt.Printf("lenght draft: %d\n", len(sh.draft))
	return sh.draft[coord.Row-1][coord.Column-1].Value
}

// GetRow returns row of sheet, row must start at 1
func (sh *Sheet) GetRow(row int) []Cell {
	if row < 1 {
		fmt.Println("row must start at 1")
		return []Cell{}
	}
	if !sh.writeAccess {
		rows, err := sh.file.GetRows(sh.name)
		if err != nil {
			fmt.Println(err)
		}
		cells := []Cell{}
		for i, value := range rows[row] {
			coords, _ := excelize.CoordinatesToCellName(i, row)
			styleID, _ := sh.file.GetCellStyle(sh.name, coords)
			cells = append(cells, Cell{Value: value, Style: RawID(styleID)})
		}
		return cells
	}
	return sh.draft[row-1]
}

// FreezeHeader freezes the headerrow
func (sh *Sheet) FreezeHeader() {
	sh.freezeHeader = true
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
		rows, _ := sh.file.GetRows(v)
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
