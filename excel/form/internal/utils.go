package internal

import (
	"fmt"
	"github.com/tidwall/gjson"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

var (
	orientation          = "landscape"
	firstPageNumber uint = 1
	blackAndWhite        = true
	color                = "FF0000"
)

// CreateSheet creates new sheet in the form
func CreateSheet(f *excelize.File, sheet string, options SheetOptions) error {
	index, err := f.NewSheet(sheet)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)
	if err = f.SetPageLayout(sheet, &excelize.PageLayoutOptions{
		Orientation:     &orientation,
		FirstPageNumber: &firstPageNumber,
		BlackAndWhite:   &blackAndWhite,
	}); err != nil {
		return err
	}

	if options.Protected {
		if err = f.ProtectSheet(sheet, &excelize.SheetProtectionOptions{
			SelectLockedCells:   true,
			SelectUnlockedCells: true,
		}); err != nil {
			return err
		}
	}

	if options.HasErrors {
		if err = f.SetSheetProps(sheet, &excelize.SheetPropsOptions{
			TabColorRGB: &color,
		}); err != nil {
			return err
		}
	}
	return nil
}

func CreateDropListSheet(f *excelize.File, sheet string) error {
	index, err := f.NewSheet(sheet)
	if err != nil {
		return err
	}

	f.SetActiveSheet(index)

	if err = f.ProtectSheet(sheet, &excelize.SheetProtectionOptions{
		SelectLockedCells: true,
	}); err != nil {
		return err
	}

	return nil
}

// GetDataFromJSON gets data field from chapters
func GetDataFromJSON(idx int, body []byte) []string {
	dataLen := int(gjson.Get(string(body), fmt.Sprintf("DOCS.%d.DATA.#", idx)).Int())
	arr := make([]string, dataLen)
	for i := 0; i < dataLen; i++ {
		arr[i] = gjson.Get(string(body), fmt.Sprintf("DOCS.%d.DATA.%d", idx, i)).String()
	}
	return arr
}

func SetColWidth(f *excelize.File, chapter string, book DOC) error {
	var colID string
	for i := 0; i < len(book.COLUMNS); i++ {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		if len(cell) > 2 {
			colID = string(cell[0]) + string(cell[1])
		} else {
			colID = string(cell[0])
		}
		colWidth := float64(book.COLUMNS[i].COLUMNWIDTHMM)
		if colWidth != 0 {
			if err := f.SetColWidth(chapter, colID, colID, colWidth); err != nil {
				return err
			}
		}
	}
	return nil
}

func SetRowHeight(f *excelize.File, chapter string, text interface{}, unitHeight int, rowID int) error {
	newLines := CountNewLines(text)
	height := float64(newLines * unitHeight)
	if err := f.SetRowHeight(chapter, rowID, height); err != nil {
		return err
	}
	return nil
}

func CountNewLines(s interface{}) int {
	maxHeight := 1
	switch v := s.(type) {
	case string:
		countNewline := strings.Count(v, "\n") + 1
		if countNewline > maxHeight {
			return countNewline
		}
	}
	return maxHeight
}

func Get2Cells(col1, row1, col2, row2 int) (string, string, error) {
	cell1, err := excelize.CoordinatesToCellName(col1, row1)
	if err != nil {
		return "", "", err
	}

	cell2, err := excelize.CoordinatesToCellName(col2, row2)
	if err != nil {
		return "", "", err
	}

	return cell1, cell2, nil
}

func ToInt(s string) int {
	num, err := strconv.Atoi(s)
	if err != nil {
		return -1
	}
	return num
}
