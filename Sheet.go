package excel

import (
	"fmt"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// Column represents a column
type Column int

func (c Column) int() int {
	return int(c)
}

func (c Column) string() string {
	str, _ := excelize.ColumnNumberToName(c.int())
	return str
}

// Row represents a Row, in which each column has a cell
type Row map[Column]*Cell

// AddStyle adds a style to row
func (row *Row) AddStyle(style Style) {
	maxCol := row.maxColumn()
	for i := 1; i < maxCol+1; i++ {
		if cell, ok := (*row)[Column(i)]; ok {
			cell.Style = style
			//(*row)[Column(i)].Style = style
		} else {
			(*row)[Column(i)] = &Cell{Value: StyleCell, Style: style}
		}
	}
}

func (row *Row) maxColumn() int {
	highest := 0
	for c := range *row {
		if c.int() > highest {
			highest = c.int()
		}
	}
	return highest
}

// Draft represents the draft of the sheet
type Draft []Row

// draft functions

func (draft *Draft) add(row Row) {
	*draft = append(*draft, row)
}

func (draft *Draft) lenght() int {
	return len(*draft)
}

func (draft *Draft) replace(index int, row Row) {
	if index == 0 {
		fmt.Println("draft index starting at 1, incrementing index")
		index++
	}
	(*draft)[index] = row
}

func (draft *Draft) value(row, column int) string {
	if row > draft.lenght() {
		fmt.Printf("row %d out of bounds, draft has only %d rows", row, draft.lenght())
		return ""
	}
	return fmt.Sprintf("%s", (*draft)[row][Column(column)].Value)
}

// CellsWithID returns a copy of all Cells with ID id
func (draft *Draft) CellsWithID(id string) []Cell {
	cells := []Cell{}
	for i, row := range *draft {
		for col, cell := range row {
			if cell.ID() != id {
				continue
			}
			if cell.coordinates == (Coordinates{}) {
				cell.coordinates = Coordinates{Column: col.int(), Row: i}
			}
			cells = append(cells, *cell)
		}
	}
	return cells
}

// Structs

// Sheet wraps the sheets of a excel file into a struct
type Sheet struct {
	file         *excelize.File
	name         string
	headerTitle  []string
	draft        Draft
	writeAccess  bool
	freezeHeader bool
}

// Get/Create Sheets
func newSheet(file *excelize.File, name string) *Sheet {
	newSheet := Sheet{file: file, name: name, headerTitle: []string{}, draft: Draft{}, writeAccess: true}
	newSheet.AddEmptyRow()
	return &newSheet
}

// Sheet retruns the sheet with the given name or creates a new one
func (excel *Excel) Sheet(name string) *Sheet {
	// Sheet exists
	for i, existingSheet := range *excel.sheets {
		if existingSheet.name == name {
			return &(*excel.sheets)[i]
		}
	}
	fmt.Printf("Creating new sheet %s\n", name)
	newSheet := newSheet(excel.file, name)
	excel.file.NewSheet(name)
	*excel.sheets = append(*excel.sheets, *newSheet)
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
func (sh *Sheet) Draft() *Draft {
	return &sh.draft
}

// GetWriteAccess populates draft with current content fo sheet and grants write access
func (sh *Sheet) GetWriteAccess() {
	if sh.writeAccess {
		fmt.Printf("write access for sheet %s already granted\n", sh.name)
		return
	}
	sh.draft = Draft{}
	sh.AddEmptyRow()
	rows, _ := sh.file.GetRows(sh.name)
	for i, row := range rows {
		newRow := Row{}
		for j, str := range row {
			styleID, _ := sh.file.GetCellStyle(sh.name, Coordinates{Row: i + 1, Column: j + 1}.String())
			newRow[Column(j+1)] = &Cell{Value: str, Style: RawID(styleID)}
		}
		sh.draft.add(newRow)
	}
	for _, cell := range sh.draft[1] {
		sh.headerTitle = append(sh.headerTitle, fmt.Sprintf("%s", cell.Value))
	}
	sh.writeAccess = true
}

// FirstSheet returns the first sheet found in the excel file
func (excel *Excel) FirstSheet() *Sheet {
	shs := excel.sheets
	return &(*shs)[0]
}

// ExtractColumnsByName extracts columns by there names from sheet
func (sh *Sheet) ExtractColumnsByName(columnTitles []string) [][]string {
	columns := []string{}
	columnMap := map[string]int{}
	for i, title := range sh.headerTitle {
		if containsString(columnTitles, title) {
			columnMap[title] = i + 1
		}
	}
	for _, title := range columnTitles {
		columnstring, err := excelize.ColumnNumberToName(columnMap[title])
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
	return len(sh.draft) + 1
}

// AddHeaderColumn adds a header column to sheet
func (sh *Sheet) AddHeaderColumn(header []string) {
	if !sh.writeAccess {
		fmt.Printf("no permission to write to sheet %s\n", sh.name)
		return
	}

	headerRow := Row{}
	for i, h := range header {
		headerRow[Column(i+1)] = &Cell{Value: h, Style: NoStyle()}
	}
	if len(sh.draft) == 1 {
		fmt.Println("Writing Header Column:")
		sh.draft.add(headerRow)
	} else {
		fmt.Println("Replacing Header Column:")
		sh.draft.replace(1, headerRow)
	}
	sh.headerTitle = header
	fmt.Println(sh.headerTitle)
}

// AddRow scanns for the next available row and inserts cells at the given indexes provided by the map
func (sh *Sheet) AddRow(row Row) {
	sh.draft.add(row)
}

// RowAtIndex returns the row at index for sheet
func (sh *Sheet) RowAtIndex(index int) Row {
	row := (*sh).draft[index]
	return row
}

// LastRowAdded returnes the row, that was last added to sheet
func (sh *Sheet) LastRowAdded() Row {
	return sh.RowAtIndex(len((*sh).draft) - 1)
}

// func (sh *Sheet) AddRow(columnCellMap map[int]Cell) {
// 	if !sh.writeAccess {
// 		fmt.Printf("no permission to write to sheet %s\n", sh.name)
// 		return
// 	}

// 	newRowIndexes := []int{}
// 	for index := range columnCellMap {
// 		newRowIndexes = append(newRowIndexes, index)
// 	}
// 	newRow := Row{}

// 	for i := 1; i != maxInt(newRowIndexes)+1; i++ {
// 		if val, ok := columnCellMap[i]; ok {
// 			val.coordinates = Coordinates{Column: i, Row: len(sh.draft) + 1}
// 			str := strings.TrimSpace(fmt.Sprintf("%s", val.Value))
// 			if str == "" {
// 				val.Value = StyleCell
// 			}
// 			newRow = append(newRow, val)
// 		} else {
// 			newRow = append(newRow, Cell{Value: DraftCell, Style: NoStyle(), coordinates: Coordinates{Column: i, Row: len(sh.draft) + 1}})
// 		}
// 	}

// 	sh.draft.add()
// }

// AddEmptyRow adds an empty row at index row
func (sh *Sheet) AddEmptyRow() {
	if !sh.writeAccess {
		fmt.Printf("no permission to write to sheet %s\n", sh.name)
		return
	}
	sh.draft.add(Row{Column(1): &Cell{Value: DraftCell, Style: NoStyle(), coordinates: Coordinates{Column: 1, Row: sh.draft.lenght() + 1}}})
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
func (sh *Sheet) GetValue(row, column int) string {
	return sh.draft.value(row, column)
}

// func (sh *Sheet) GetValue(coord Coordinates) interface{} {
// 	if !sh.writeAccess {
// 		value, err := sh.file.GetCellValue(sh.name, coord.String())
// 		if err != nil {
// 			fmt.Println(err)
// 		}
// 		return value
// 	}
// 	fmt.Printf("lenght draft: %d\n", len(sh.draft))
// 	return sh.draft[coord.Row-1][coord.Column-1].Value
// }

// GetRow returns row of sheet, row must start at 1
func (sh *Sheet) GetRow(row int) Row {
	newRow := Row{}
	if row < 1 {
		fmt.Println("row must start at 1")
		return newRow
	}
	if !sh.writeAccess {
		rows, err := sh.file.GetRows(sh.name)
		if err != nil {
			fmt.Println(err)
		}
		for i, value := range rows[row-1] {
			coords, _ := excelize.CoordinatesToCellName(i, row)
			styleID, _ := sh.file.GetCellStyle(sh.name, coords)
			newRow[Column(i+1)] = &Cell{Value: value, Style: RawID(styleID)}
		}
		return newRow
	}
	return sh.draft[row]
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
	return sh.headerTitle
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
