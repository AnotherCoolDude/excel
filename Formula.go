package excel

import (
	"fmt"
)

// Formula wraps the coords that will be contained by the formula in a struct
type Formula struct {
	Coords []Coordinates
}

// Add adds the coords
func (formula *Formula) Add() string {
	if len(formula.Coords) == 0 {
		return "0"
	}
	str := "="
	for _, c := range formula.Coords {
		str += c.ToString()
		if c != formula.Coords[len(formula.Coords)-1] {
			str += "+"
		}
	}
	return str
}

// Substract substracts the provided coords. The minuend is defined by the function in parameter
func (formula *Formula) Substract(fn func(coords []Coordinates) Coordinates) string {
	if len(formula.Coords) == 0 {
		return "0"
	}
	min := fn(formula.Coords)
	str := fmt.Sprintf("=%s", min.ToString())
	for _, sub := range formula.Coords {
		if sub.ToString() == min.ToString() {
			continue
		}
		str += fmt.Sprintf("-%s", sub.ToString())
	}
	return str
}

// Raw provides the coords and expects a excel-ready string
func (formula *Formula) Raw(fn func(coords []Coordinates) string) string {
	if len(formula.Coords) == 0 {
		return "0"
	}
	return fn(formula.Coords)
}
