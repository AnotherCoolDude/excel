package excel

// Formula wraps the coords that will be contained by the formula in a struct
type Formula struct {
	Coords []Coordinates
}

// Add adds the coords
func (formula *Formula) Add() string {
	str := "="
	for idx, c := range formula.Coords {
		str += c.ToString()

	}
	return str
}
