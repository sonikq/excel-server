package internal

import (
	"errors"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Style struct {
	Alignment  Alignment
	Font       Font
	Fill       Fill
	Protection Protection
}

type Alignment struct {
	Horizontal string
	Vertical   string
	Indent     int
}
type Font struct {
	Bold   bool
	Italic bool
	Family string
	Size   float64
}

type Fill struct {
	Color string
}

type Protection struct {
	Hidden bool
	Locked bool
}

func setCommonStyle(cellStyle *Style, book DOC, f *excelize.File) (int, error) {
	if cellStyle.Font.Family == "" {
		cellStyle.Font.Family = book.COMMON.STYLE.FONT
	}

	if cellStyle.Alignment.Vertical == "" {
		cellStyle.Alignment.Vertical = book.COMMON.STYLE.VERTALIGNMENT
	}

	if cellStyle.Alignment.Horizontal == "" {
		cellStyle.Alignment.Horizontal = book.COMMON.STYLE.HORALIGNMENT
	}

	if cellStyle.Font.Size == 0 {
		cellStyle.Font.Size = book.COMMON.STYLE.SIZE
	}

	unitStyle := excelize.Style{
		Alignment: &excelize.Alignment{
			Vertical:   strings.ToLower(cellStyle.Alignment.Vertical),
			Horizontal: strings.ToLower(cellStyle.Alignment.Horizontal),
			Indent:     cellStyle.Alignment.Indent,
			WrapText:   true,
		},
		Font: &excelize.Font{
			Family: cellStyle.Font.Family,
			Bold:   cellStyle.Font.Bold,
			Size:   cellStyle.Font.Size,
			Italic: cellStyle.Font.Italic,
		},
		Border: []excelize.Border{
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
		},
		Protection: &excelize.Protection{
			Locked: cellStyle.Protection.Locked,
		},
	}

	if cellStyle.Fill.Color != "" {
		unitStyle.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{cellStyle.Fill.Color},
			Pattern: 1,
		}
	}

	style, err := f.NewStyle(&unitStyle)
	if err != nil {
		return -1, err
	}

	return style, err
}

func setFooterStyle(cellStyle *Style, book DOC, f *excelize.File) (int, error) {
	//if cellStyle.Font.Family == "" {
	//	cellStyle.Font.Family = book.COMMON.STYLE.FONT
	//}
	//
	//if cellStyle.Alignment.Vertical == "" {
	//	cellStyle.Alignment.Vertical = book.COMMON.STYLE.VERTALIGNMENT
	//}
	//
	//if cellStyle.Alignment.Horizontal == "" {
	//	cellStyle.Alignment.Horizontal = book.COMMON.STYLE.HORALIGNMENT
	//}
	//
	//if cellStyle.Font.Size == 0 {
	//	cellStyle.Font.Size = book.COMMON.STYLE.SIZE
	//}

	cellStyle.Font.Family = book.COMMON.STYLE.FONT
	cellStyle.Alignment.Vertical = book.COMMON.STYLE.VERTALIGNMENT
	cellStyle.Alignment.Horizontal = book.COMMON.STYLE.HORALIGNMENT
	cellStyle.Font.Size = book.COMMON.STYLE.SIZE

	unitStyle := excelize.Style{
		Alignment: &excelize.Alignment{
			Vertical:   strings.ToLower(cellStyle.Alignment.Vertical),
			Horizontal: strings.ToLower(cellStyle.Alignment.Horizontal),
			Indent:     cellStyle.Alignment.Indent,
			WrapText:   true,
		},
		Font: &excelize.Font{
			Family: cellStyle.Font.Family,
			Bold:   cellStyle.Font.Bold,
			Size:   cellStyle.Font.Size,
			Italic: cellStyle.Font.Italic,
		},
		Protection: &excelize.Protection{
			Locked: cellStyle.Protection.Locked,
		},
	}

	if cellStyle.Fill.Color != "" {
		unitStyle.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{cellStyle.Fill.Color},
			Pattern: 1,
		}
	}

	style, err := f.NewStyle(&unitStyle)
	if err != nil {
		return -1, err
	}

	return style, err
}

func setCellStyles(f *excelize.File, chapter string, book DOC, dataRowMap map[int]int) error {

	for i, style := range book.STYLES {
		if len(style.GRID) != 4 {
			return errors.New("grid is invalid")
		}

		for j := style.GRID[1] - 1; j <= style.GRID[3]-1; j++ {
			styleID, err := stylize(f, book, i, j)
			if err != nil {
				return err
			}

			rowID := dataRowMap[style.GRID[0]]
			if rowID == 0 {
				continue
			}
			cell1, err := excelize.CoordinatesToCellName(j+1, rowID)
			if err != nil {
				return err
			}

			rowID = dataRowMap[style.GRID[2]]
			cell2, err := excelize.CoordinatesToCellName(j+1, rowID)
			if err != nil {
				return err
			}

			if err = f.SetCellStyle(chapter, cell1, cell2, styleID); err != nil {
				return err
			}
		}
	}

	return nil
}

func stylize(f *excelize.File, book DOC, i int, j int) (int, error) {
	var cellStyle Style

	if book.STYLES[i].CELL_STYLE.FONT == "" {
		if book.COLUMNS[j].STYLE.FONT == "" {
			cellStyle.Font.Family = book.COMMON.STYLE.FONT
		} else {
			cellStyle.Font.Family = book.COLUMNS[j].STYLE.FONT
		}
	} else {
		cellStyle.Font.Family = book.STYLES[i].CELL_STYLE.FONT
	}

	if book.STYLES[i].CELL_STYLE.SIZE == 0 {
		if book.COLUMNS[j].STYLE.SIZE == 0 {
			cellStyle.Font.Size = book.COMMON.STYLE.SIZE
		} else {
			cellStyle.Font.Size = book.COLUMNS[j].STYLE.SIZE
		}
	} else {
		cellStyle.Font.Size = book.STYLES[i].CELL_STYLE.SIZE
	}

	if book.STYLES[i].CELL_STYLE.VERTALIGNMENT == "" {
		if book.COLUMNS[j].STYLE.VERTALIGNMENT == "" {
			cellStyle.Alignment.Vertical = book.COMMON.STYLE.VERTALIGNMENT
		} else {
			cellStyle.Alignment.Vertical = book.COLUMNS[j].STYLE.VERTALIGNMENT
		}
	} else {
		cellStyle.Alignment.Vertical = book.STYLES[i].CELL_STYLE.VERTALIGNMENT
	}

	if book.STYLES[i].CELL_STYLE.HORALIGNMENT == "" {
		if book.COLUMNS[j].STYLE.HORALIGNMENT == "" {
			cellStyle.Alignment.Horizontal = book.COMMON.STYLE.HORALIGNMENT
		} else {
			cellStyle.Alignment.Horizontal = book.COLUMNS[j].STYLE.HORALIGNMENT
		}
	} else {
		cellStyle.Alignment.Horizontal = book.STYLES[i].CELL_STYLE.HORALIGNMENT
	}

	if book.STYLES[i].CELL_STYLE.INDENT == 0 {
		if book.COLUMNS[j].STYLE.INDENT == 0 {
			cellStyle.Alignment.Indent = book.COMMON.STYLE.INDENT
		} else {
			cellStyle.Alignment.Indent = book.COLUMNS[j].STYLE.INDENT
		}
	} else {
		cellStyle.Alignment.Indent = book.STYLES[i].CELL_STYLE.INDENT
	}

	if !book.STYLES[i].CELL_STYLE.FONT_STYLE.BOLD {
		if !book.COLUMNS[j].STYLE.FONT_STYLE.BOLD {
			cellStyle.Font.Bold = false
		} else {
			cellStyle.Font.Bold = book.COLUMNS[j].STYLE.FONT_STYLE.BOLD
		}
	} else {
		cellStyle.Font.Bold = book.STYLES[i].CELL_STYLE.FONT_STYLE.BOLD
	}

	cellStyle.Protection.Locked = book.STYLES[i].PROTECTED

	unitStyle := excelize.Style{
		Alignment: &excelize.Alignment{
			Vertical:   strings.ToLower(cellStyle.Alignment.Vertical),
			Horizontal: strings.ToLower(cellStyle.Alignment.Horizontal),
			Indent:     cellStyle.Alignment.Indent,
			WrapText:   true,
		},
		Font: &excelize.Font{
			Family: cellStyle.Font.Family,
			Bold:   cellStyle.Font.Bold,
			Size:   cellStyle.Font.Size,
			Italic: cellStyle.Font.Italic,
		},
		Border: []excelize.Border{
			{
				Type:  "top",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "left",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "000000",
				Style: 1,
			},
			{
				Type:  "right",
				Color: "000000",
				Style: 1,
			},
		},
		Protection: &excelize.Protection{
			Locked: cellStyle.Protection.Locked,
		},
	}

	if cellStyle.Fill.Color != "" {
		unitStyle.Fill = excelize.Fill{
			Type:    "pattern",
			Color:   []string{cellStyle.Fill.Color},
			Pattern: 1,
		}
	}

	styleID, err := f.NewStyle(&unitStyle)
	if err != nil {
		return -1, err
	}

	return styleID, nil
}
