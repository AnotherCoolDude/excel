package excel

import (
	"github.com/360EntSecGroup-Skylar/excelize"
)

// Coordinates wraps coordinates in a struct
type Coordinates struct {
	Row, Column int
}

// ToString returns the coordinates as excelformatted string
func (c Coordinates) ToString() string {
	if c.Row == 0 {
		c.Row = 1
	}
	return excelize.MustCoordinatesToCellName(c.Column, c.Row)
}
