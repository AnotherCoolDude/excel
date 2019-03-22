package excel

// Insertable defines Methods for structs to be insertable in a excelfile
type Insertable interface {
	Insert(sh *Sheet)
}

// Add inserts a insertable struct into a given file.
func (sh *Sheet) Add(data Insertable) {
	data.Insert(sh)
}
