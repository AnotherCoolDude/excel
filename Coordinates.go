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
	if c.Row == 0 {
		c.Row = 1
	}
	return fmt.Sprintf("%s%d", excelize.ToAlphaString(c.Column), c.Row)
}
