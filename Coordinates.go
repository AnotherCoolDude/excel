package excel

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
)

// Coordinates wraps coordinates in a struct
type Coordinates struct {
	Row, Column int
}

// ToString returns the coordinates as excelformatted string
func (c Coordinates) String() string {
	if c.Row == 0 || c.Column == 0 {
		fmt.Printf("Coordinates: %d, %d -> 0 is not allowed, excel starts at 1\n", c.Column, c.Row)
		return ""
	}
	str, err := excelize.CoordinatesToCellName(c.Column, c.Row)
	if err != nil {
		fmt.Println(err)
	}
	return str
}

// StringWithReference returns the coordinates as excelformatted string, which references to another sheet
func (c Coordinates) StringWithReference(sheet string) string {
	if sheet == "" {
		return c.String()
	}
	str := c.String()
	ref := fmt.Sprintf("'%s'!", sheet)
	return ref + str
}
