package excel

// // CellMap stores Cells identified by a string
// type CellMap map[string][]Coordinates

// // Context provides a global storage for cellMaps
// type Context struct {
// 	header   *Header
// 	cellMaps *map[string]CellMap
// }

// // NewContext returns a new context struct
// func NewContext(header *Header) Context {
// 	return Context{
// 		header:   header,
// 		cellMaps: &map[string]CellMap{},
// 	}
// }

// // AddCellMap adds a cellMap to context
// func (ctx *Context) AddCellMap(cellMap *CellMap, id string) {
// 	(*ctx.cellMaps)[id] = *cellMap
// }

// // AddFromCurrentRow adds cells from current row and headerColumns to cellMap with id mapID
// func (ctx *Context) AddFromCurrentRow(sh *Sheet, mapID string, HeaderColumnList []HeaderColumn) {
// 	cellMap := (*ctx.cellMaps)[mapID]
// 	for _, hdr := range HeaderColumnList {
// 		cellMap[ctx.header.TitleForColumn(hdr)] = append(cellMap[ctx.header.TitleForColumn(hdr)], Coordinates{Row: sh.CurrentRow(), Column: int(hdr)})
// 	}
// 	(*ctx.cellMaps)[mapID] = cellMap
// }

// // AddFromRow adds cells from row and  headerColumns to cellMap with id mapID
// func (ctx *Context) AddFromRow(sh *Sheet, mapID string, row int, HeaderColumnList []HeaderColumn) {
// 	cellMap := (*ctx.cellMaps)[mapID]
// 	for _, hdr := range HeaderColumnList {
// 		cellMap[ctx.header.TitleForColumn(hdr)] = append(cellMap[ctx.header.TitleForColumn(hdr)], Coordinates{Row: row, Column: int(hdr)})
// 	}
// }

// // Formula returns a Formula struct for cellMap with mapID and HeaderColumn
// func (ctx *Context) Formula(mapID string, hdr HeaderColumn) *Formula {
// 	cellMap := (*ctx.cellMaps)[mapID]
// 	coords := cellMap[ctx.header.TitleForColumn(hdr)]
// 	return &Formula{Coords: &coords}
// }

// // ResetCellMap resets cellMap with mapID and returns the cellMap, thats being resetted
// func (ctx *Context) ResetCellMap(mapID string) CellMap {
// 	cellMap := (*ctx.cellMaps)[mapID]
// 	(*ctx.cellMaps)[mapID] = CellMap{}
// 	return cellMap
// }
