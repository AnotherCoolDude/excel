package excel

import "fmt"

// Insertable defines Methods for structs to be insertable in a excelfile
type Insertable interface {
	Columns() []string
	Insert(sh *Sheet)
}

// Add inserts a insertable struct into a given file.
func (sh *Sheet) Add(data Insertable) {
	if sh.isEmpty() {
		fmt.Println("file is empty, adding header")
		headerCoords := Coordinates{Row: 0, Column: 0}
		if len(data.Columns()) == 0 {
			fmt.Printf("provide at least one struct, that satisfies Insertable and returns not an empty slice of strings in the Columns method")
			return
		}
		if sh.draftMode {
			headerCells := []Cell{}
			for _, header := range data.Columns() {
				headerCells = append(headerCells, Cell{Value: header, Style: NoStyle()})
			}
			sh.draft = append(sh.draft, headerCells)
			sh.columns = data.Columns()
		} else {
			for _, col := range data.Columns() {
				fmt.Printf("writing header %s at %s\n", col, headerCoords.ToString())
				sh.file.SetCellStr(sh.name, headerCoords.ToString(), col)
				headerCoords.Column = headerCoords.Column + 1
			}
		}
	}
	data.Insert(sh)
}
