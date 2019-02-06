package excel

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
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
		sheets = append(sheets, Sheet{file: eFile, name: sheetname, columns: []string{}})
	} else {
		eFile, err = excelize.OpenFile(path)
		sheetMap := eFile.GetSheetMap()
		for _, name := range sheetMap {
			header := eFile.GetRows(name)[0]
			sheets = append(sheets, Sheet{file: eFile, name: name, columns: header})
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
	excel.file.SaveAs(path)
}
