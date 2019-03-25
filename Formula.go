package excel

import (
	"fmt"
)

// Formula wraps the coords that will be contained by the formula in a struct
type Formula struct {
	Coords *[]Coordinates
	sheet  string
}

// Reference makes the formula reference to another sheet
func (formula *Formula) Reference(sheet string) *Formula {
	formula.sheet = sheet
	return formula
}

// Sum sums up the provided coords
func (formula *Formula) Sum() string {
	if len(*formula.Coords) == 0 {
		return "0"
	}

	lowest := (*formula.Coords)[0]
	highest := (*formula.Coords)[0]

	for _, coord := range *formula.Coords {
		if coord.Row <= lowest.Row && coord.Column <= lowest.Column {
			lowest = coord
		}
		if coord.Row >= highest.Row && coord.Column >= highest.Column {
			highest = coord
		}
	}

	return fmt.Sprintf("=SUMME(%s:%s)", lowest.StringWithReference(formula.sheet), highest.StringWithReference(formula.sheet))
}

// Add adds the coords
func (formula *Formula) Add() string {
	if len(*formula.Coords) == 0 {
		return "0"
	}

	str := "="
	for _, c := range *formula.Coords {
		str += c.StringWithReference(formula.sheet)
		if c != (*formula.Coords)[len(*formula.Coords)-1] {
			str += "+"
		}
	}
	return str

}

// Substract substracts the provided coords. The minuend is defined by the function in parameter
func (formula *Formula) Substract(fn func(coords []Coordinates) Coordinates) string {
	if len(*formula.Coords) == 0 {
		return "0"
	}
	min := fn(*formula.Coords)
	str := fmt.Sprintf("=%s", min.StringWithReference(formula.sheet))
	for _, sub := range *formula.Coords {
		if sub.StringWithReference(formula.sheet) == min.StringWithReference(formula.sheet) {
			continue
		}
		str += fmt.Sprintf("-%s", sub.StringWithReference(formula.sheet))
	}
	return str
}

// Raw provides the coords and expects a excel-ready string
func (formula *Formula) Raw(fn func(coords []Coordinates) string) string {
	if len(*formula.Coords) == 0 {
		return "0"
	}
	return fn(*formula.Coords)
}
