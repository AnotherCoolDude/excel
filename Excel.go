package excel

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/schollz/progressbar"
)

const (
	draftCell = "DRAFT_CELL"
)

// Excel wraps the excelize package
type Excel struct {
	file   *excelize.File
	sheets *[]Sheet
}

// File opens/creates a Excel file. If newly created, names the first sheet after sheetname
func File(path string, sheetname string) *Excel {
	var eFile *excelize.File
	var sheets []Sheet
	if _, err := os.Stat(path); os.IsNotExist(err) {
		fmt.Println("file not existing, creating new...")
		eFile = excelize.NewFile()
		sheetIndex := eFile.GetActiveSheetIndex()
		oldName := eFile.GetSheetName(sheetIndex)
		eFile.SetSheetName(oldName, sheetname)
		sheets = append(sheets, Sheet{file: eFile, name: sheetname, columns: []string{}, draft: [][]Cell{}, writeAccess: true})
	} else {
		fmt.Printf("found file at path %s\n", path)
		eFile, err = excelize.OpenFile(path)
		sheetMap := eFile.GetSheetMap()
		for _, name := range sheetMap {
			rows, _ := eFile.GetRows(name)
			header := rows[0]
			sheets = append(sheets, Sheet{file: eFile, name: name, columns: header, writeAccess: false})
		}
		if err != nil {
			fmt.Printf("couldn't open file at path\n%s\nerr: %s", path, err)
		}
	}
	return &Excel{
		file:   eFile,
		sheets: &sheets,
	}
}

// Save saves the Excelfile to the provided path
func (excel *Excel) Save(path string) {
	for _, sheet := range *excel.sheets {
		if !sheet.writeAccess {
			fmt.Printf("WARNING: didn't save. No write access for sheet %s\n", sheet.name)
			continue
		}
		bar := progressbar.New(len(sheet.draft))
		sheet.clearSheet()
		currentCoords := Coordinates{Row: 0, Column: 0}
		for i, row := range sheet.draft {
			for j, cell := range row {
				bar.Add(1)
				if cell.Value == draftCell {
					continue
				}
				currentCoords.Row = i + 1
				currentCoords.Column = j + 1
				excel.file.SetCellValue(sheet.name, currentCoords.ToString(), cell.Value)

				if isRaw, id := cell.Style.RawID(); isRaw {
					excel.file.SetCellStyle(sheet.name, currentCoords.ToString(), currentCoords.ToString(), id)
					continue
				}

				styleString := cell.Style.toString()
				if styleString == "" {
					continue
				}
				st, err := excel.file.NewStyle(styleString)
				if err != nil {
					fmt.Println(styleString)
					fmt.Println(err)
				}
				excel.file.SetCellStyle(sheet.name, currentCoords.ToString(), currentCoords.ToString(), st)
			}
		}

	}

	excel.file.SaveAs(path)
	println()
}
