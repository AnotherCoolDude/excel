package excel

import "fmt"

// Cell wraps a cell's value and its style in a struct
type Cell struct {
	Value       interface{}
	Style       Style
	coordinates Coordinates
}

// Coordinates returns the coordinates associated with cell
func (c *Cell) Coordinates() Coordinates {
	if c.coordinates == (Coordinates{}) {
		fmt.Println("Coordinates for Cell are not yet initialized, returning empty stuct")
		return Coordinates{}
	}
	return c.coordinates
}

// ChangeStyle changes Style of cell
func (c *Cell) ChangeStyle(style Style) *Cell {
	c.Style = style
	return c
}
