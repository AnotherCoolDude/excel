package excel

import "errors"

// Cell wraps a cell's value and its style in a struct
type Cell struct {
	Value       interface{}
	Style       Style
	coordinates Coordinates
}

// Coordinates returns the coordinates associated with cell
func (c *Cell) Coordinates() (Coordinates, error) {
	if c.coordinates == (Coordinates{}) {
		return Coordinates{}, errors.New("Coordinates for Cell are not yet initialized, returning empty stuct")
	}
	return c.coordinates, nil
}

// ChangeStyle changes Style of cell
func (c *Cell) ChangeStyle(style Style) *Cell {
	c.Style = style
	return c
}
