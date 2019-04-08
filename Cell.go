package excel

import "fmt"

// Cell wraps a cell's value and its style in a struct
type Cell struct {
	Value       interface{}
	Style       Style
	id          string
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

// SetID adds an id to Cell
func (c *Cell) SetID(value string) *Cell {
	c.id = value
	return c
}

// ID returns the id of cell
func (c *Cell) ID() string {
	return c.id
}

// NewCell returns a new Cell with no specific style
func NewCell(value string) *Cell {
	return &Cell{Value: value, Style: NoStyle()}
}

// NewEuroCell returns a new Cell with Euro formatting
func NewEuroCell(value float32) *Cell {
	return NewCell(fmt.Sprintf("%.2f", value)).ChangeStyle(EuroStyle())
}

// ChangeStyle changes Style of cell
func (c *Cell) ChangeStyle(style Style) *Cell {
	c.Style = style
	return c
}

// HasValue returns true, if cell has a value
func (c *Cell) HasValue() bool {
	if c.Value == DraftCell || c.Value == StyleCell {
		return false
	}
	return true
}
