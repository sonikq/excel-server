package internal

import "github.com/xuri/excelize/v2"

func setTitleStyle(f *excelize.File, fontSize float64, isBold bool, isBordered bool) (int, error) {
	var borders []excelize.Border

	if isBordered {
		borders = append(borders, []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1}}...)
	}

	titleStyle := excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
			WrapText:   true,
		},
		Font: &excelize.Font{
			Bold:   isBold,
			Size:   fontSize,
			Family: "Times New Roman",
		},
		Border: borders,
	}

	styleID, err := f.NewStyle(&titleStyle)
	if err != nil {
		return -1, err
	}

	return styleID, err
}

func setTitleResponsesStyle(f *excelize.File, fontSize float64, indentVal int, isBold bool, isTopBordered, isBottomBordered, isLeftBordered, isRightBordered bool) (int, error) {
	var borders []excelize.Border

	if isTopBordered {
		borders = append(borders, []excelize.Border{
			{Type: "top", Color: "000000", Style: 1}}...)
	}
	if isBottomBordered {
		borders = append(borders, []excelize.Border{
			{Type: "bottom", Color: "000000", Style: 1}}...)
	}
	if isLeftBordered {
		borders = append(borders, []excelize.Border{
			{Type: "left", Color: "000000", Style: 1}}...)
	}
	if isRightBordered {
		borders = append(borders, []excelize.Border{
			{Type: "right", Color: "000000", Style: 1}}...)
	}

	titleStyle := excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
			Indent:     indentVal,
			WrapText:   true,
		},
		Font: &excelize.Font{
			Bold:   isBold,
			Size:   fontSize,
			Family: "Times New Roman",
		},
		Border: borders,
	}

	styleID, err := f.NewStyle(&titleStyle)
	if err != nil {
		return -1, err
	}

	return styleID, err
}

func setTitleHeaderFooterStyle(f *excelize.File, fontSize float64, isBold bool, isBordered bool) (int, error) {
	var borders []excelize.Border

	if isBordered {
		borders = append(borders, []excelize.Border{
			{Type: "top", Color: "000000", Style: 1},
			{Type: "left", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1}}...)
	}

	footerStyle := excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
			WrapText:   true,
		},
		Font: &excelize.Font{
			Bold:   isBold,
			Size:   fontSize,
			Family: "Times New Roman",
		},
		Border: borders,
	}

	styleID, err := f.NewStyle(&footerStyle)
	if err != nil {
		return -1, err
	}

	return styleID, err
}
