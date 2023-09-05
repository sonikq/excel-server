package internal

import "github.com/xuri/excelize/v2"

const (
	errFirstHeader  = "Раздел"
	errSecondHeader = "Ячейка"
	errThirdHeader  = "Введенное значение"
)

func (f *Form) printErrors() error {

	sheetOpts := SheetOptions{
		Protected: false,
		HasErrors: true,
	}

	if err := CreateSheet(f.File, f.ErrorsSheet, sheetOpts); err != nil {
		return err
	}

	if err := printErrorsHeader(f.File, f.ErrorsSheet); err != nil {
		return err
	}

	for i, uploadErr := range f.Errors {
		if err := printErrData(f.File, f.ErrorsSheet, uploadErr, f.ErrorsRowID+i); err != nil {
			return err
		}
	}

	return nil
}

func printErrorsHeader(f *excelize.File, sheet string) error {

	headers := []string{errFirstHeader, errSecondHeader, errThirdHeader}
	colWidths := []float64{25, 15, 40}

	for i := 0; i <= 2; i++ {
		cell1, err := excelize.CoordinatesToCellName(i+1, 1)
		if err != nil {
			return err
		}

		style, err := setTitleStyle(f, 10, true, true)
		if err != nil {
			return err
		}

		cell := Cell{
			Data:    headers[i],
			Cell1:   cell1,
			Cell2:   cell1,
			StyleID: style,
		}

		if err = printAnyCell(f, sheet, cell); err != nil {
			return err
		}

		if err = f.SetColWidth(sheet, string(cell1[0]), string(cell1[0]), colWidths[i]); err != nil {
			return err
		}
	}

	return nil
}

func printErrData(f *excelize.File, sheet string, data []interface{}, errRowID int) error {
	for i := 0; i < 3; i++ {
		cell1, err := excelize.CoordinatesToCellName(i+1, errRowID)
		if err != nil {
			return err
		}

		style, err := setTitleStyle(f, 10, false, true)
		if err != nil {
			return err
		}

		cell := Cell{
			Data:    data[i],
			Cell1:   cell1,
			Cell2:   cell1,
			StyleID: style,
		}

		if err = printAnyCell(f, sheet, cell); err != nil {
			return err
		}
	}
	return nil
}
