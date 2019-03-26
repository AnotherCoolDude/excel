package excel

// import (
// 	"fmt"
// )

// // HeaderColumn defines the title of a column
// type HeaderColumn int

// // Header wraps the headercolumn in a struct
// type Header struct {
// 	headerMap map[HeaderColumn]string
// }

// func newHeaderColumn(title []string) Header {
// 	header := Header{}
// 	header.headerMap = map[HeaderColumn]string{}
// 	for i, t := range title {
// 		header.headerMap[HeaderColumn(i+1)] = t
// 	}
// 	return header
// }

// // Title returns the title of headerolumn
// func (hdr *Header) Title() []string {
// 	title := []string{}
// 	for _, header := range hdr.headerMap {
// 		title = append(title, header)
// 	}
// 	return title
// }

// // TitleForColumn returns the title for column
// func (hdr *Header) TitleForColumn(column HeaderColumn) string {
// 	for k, v := range hdr.headerMap {
// 		if k == column {
// 			return v
// 		}
// 	}
// 	fmt.Printf("Couldn't find title for column %d\n", column)
// 	return ""
// }

// // ColumnForTitle returns the HeaderColumn for title
// func (hdr *Header) ColumnForTitle(title string) HeaderColumn {
// 	for k, v := range hdr.headerMap {
// 		if v == title {
// 			return k
// 		}
// 	}
// 	return HeaderColumn(0)
// }
