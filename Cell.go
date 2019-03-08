package excel

// Cell wraps a cell's value and its style in a struct
type Cell struct {
	Value interface{}
	Style Style
}
