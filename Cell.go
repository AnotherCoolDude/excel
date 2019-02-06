package excel

// Cell wraps a cell's value and it's style in a struct
type Cell struct {
	Value interface{}
	Style Style
}
