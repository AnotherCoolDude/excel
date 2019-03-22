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
func (c Coordinates) ToString() string {
	if c.Row == 0 || c.Column == 0 {
		fmt.Printf("Coordinates: %v -> 0 is not allowed, excel starts at 1\n", c)
		return ""
	}
	return excelize.MustCoordinatesToCellName(c.Column, c.Row)
}
