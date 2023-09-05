package internal

import (
	"github.com/xuri/excelize/v2"
)

type Cell struct {
	Data    interface{}
	Cell1   string
	Cell2   string
	StyleID int
}

// printAnyCell prints anything
func printAnyCell(f *excelize.File, sheet string, cell Cell) error {
	if err := f.SetCellValue(sheet, cell.Cell1, cell.Data); err != nil {
		return err
	}

	if err := f.MergeCell(sheet, cell.Cell1, cell.Cell2); err != nil {
		return err
	}

	if err := f.SetCellStyle(sheet, cell.Cell1, cell.Cell2, cell.StyleID); err != nil {
		return err
	}

	return nil
}

// setRowBreaksHeight sets heights of title page's row breaks
func (title *title) setRowBreaksHeight(lenResp int) error {
	rowBreaks := [][]int{
		{3, 2}, {4, 4}, {5, 6},
		{4, 8}, {3, 11}, {4, 12},
		{4, 14 + 2*lenResp + 1},
	}

	if lenResp == 1 {
		rowBreaks = append(rowBreaks, []int{12, 14 + 2*lenResp})
	} else {
		rowBreaks = append(rowBreaks, []int{3, 14 + 2*lenResp})
	}

	for _, rowBreak := range rowBreaks {
		if err := SetRowHeight(title.file, title.sheet, "", rowBreak[0], rowBreak[1]); err != nil {
			return err
		}
	}

	return nil
}
