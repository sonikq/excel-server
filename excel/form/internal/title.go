package internal

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type title struct {
	sheet     string
	nextRowID int
	file      *excelize.File
}

func newTitle(tp *TitlePage, file *excelize.File) *title {
	return &title{
		sheet:     "Титул " + tp.FormName,
		nextRowID: 14,
		file:      file,
	}
}

// PrintTitle печатает титульный лист
func PrintTitle(f *Form, tp *TitlePage) error {
	t := newTitle(tp, f.File)

	sheetOpts := SheetOptions{
		Protected: false,
		HasErrors: false,
	}

	if err := CreateSheet(t.file, t.sheet, sheetOpts); err != nil {
		return err
	}

	if err := t.setRowBreaksHeight(len(tp.Responses)); err != nil {
		return err
	}

	if err := t.printStaticHeaders(); err != nil {
		return err
	}

	if err := t.printTitle(tp.Title); err != nil {
		return err
	}

	if err := t.printStateByDate(tp.StateByDate); err != nil {
		return err
	}

	if err := t.printFormName(tp.FormName); err != nil {
		return err
	}

	if err := t.printOrder(tp.OrderDate, tp.OrderNum); err != nil {
		return err
	}

	if err := t.printReportType(tp.ReportType); err != nil {
		return err
	}

	for _, r := range tp.Responses {

		if err := t.printResponses(r.ORGResponder, r.ORGReviewers, tp); err != nil {
			return err
		}

		if err := t.printDeadlines(r.Deadline); err != nil {
			return err
		}
		t.nextRowID += 2
	}

	t.nextRowID += 2

	if err := t.printFooter(tp.Organization, tp.OKUD, tp.OKPO); err != nil {
		return err
	}
	return nil
}

// printStaticHeaders includes functions that print static headers
// and sets columns` widths
func (title *title) printStaticHeaders() error {
	if err := title.printFirstHeader(); err != nil {
		return err
	}

	if err := title.printSecondHeader(); err != nil {
		return err
	}

	if err := title.printThirdHeader(); err != nil {
		return err
	}

	if err := title.printFourthHeader(); err != nil {
		return err
	}

	if err := title.printResponseHeaders(); err != nil {
		return err
	}

	if err := title.setColWidths(); err != nil {
		return err
	}

	return nil
}

func (title *title) printFooter(org, okud, okpo string) error {
	if err := title.printTitleFooterFirstHeader(org); err != nil {
		return err
	}

	if err := title.printTitleFooterSecondHeader(); err != nil {
		return err
	}

	if err := title.printTitleFooterThirdHeader(); err != nil {
		return err
	}

	if err := title.printTitleFooterFourthHeader(); err != nil {
		return err
	}

	if err := title.printTitleFooterPreHeader(); err != nil {
		return err
	}

	if err := title.printTitleFooterLabel(); err != nil {
		return err
	}

	if err := title.printTitleFooterData(okud, okpo); err != nil {
		return err
	}

	return nil
}

// printTitle prints title
func (title *title) printTitle(titleName string) error {
	cell1, cell2, _ := Get2Cells(5, 9, 11, 9)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    titleName,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	if err = SetRowHeight(title.file, title.sheet, titleName, unitHeightForMO, 9); err != nil {
		return err
	}

	return nil
}

// printStateByDate prints state by date
func (title *title) printStateByDate(stateByDate string) error {
	cell1, cell2, _ := Get2Cells(5, 10, 11, 10)

	style, err := setTitleStyle(title.file, 11, true, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    stateByDate,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	if err = SetRowHeight(title.file, title.sheet, stateByDate, unitTitleHeight, 10); err != nil {
		return err
	}

	return nil
}

// printFormName prints form name
func (title *title) printFormName(formName string) error {
	cell1, cell2, _ := Get2Cells(11, 13, 14, 13)

	style, err := setTitleStyle(title.file, 10, true, true)
	if err != nil {
		return err
	}

	fName := fmt.Sprintf(FormNameLayout, formName)

	cell := Cell{
		Data:    fName,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	if err = SetRowHeight(title.file, title.sheet, formName, unitTitleHeight, 13); err != nil {
		return err
	}

	return nil
}

func (title *title) printOrder(orderDate string, orderNum string) error {
	cell1, cell2, _ := Get2Cells(10, 14, 15, 15)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	order := fmt.Sprintf(OrderLayout, orderDate, orderNum)

	cell := Cell{
		Data:    order,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	return nil
}

// printReportType prints report type
func (title *title) printReportType(reportType string) error {
	cell1, cell2, _ := Get2Cells(11, 16, 14, 16)

	style, err := setTitleStyle(title.file, 10, true, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    reportType,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	return nil
}

func (title *title) printResponseHeaders() error {
	cell1, cell2, _ := Get2Cells(1, 13, 6, 13)
	cell3, cell4, _ := Get2Cells(7, 13, 8, 13)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	firstHeader := Cell{
		Data:    titleResponseFirstHeader,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, firstHeader); err != nil {
		return err
	}

	secondHeader := Cell{
		Data:    titleResponseSecondHeader,
		Cell1:   cell3,
		Cell2:   cell4,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, secondHeader); err != nil {
		return err
	}

	return nil
}

// printResponses prints responder and list of reviewers
func (title *title) printResponses(responder string, reviewers []string, t *TitlePage) error {
	responderCell1, responderCell2, _ := Get2Cells(1, title.nextRowID, 6, title.nextRowID)

	style, err := setTitleResponsesStyle(title.file, 10, 0, false, false, false, true, true)
	if err != nil {
		return err
	}

	responderData := Cell{
		Data:    responder,
		Cell1:   responderCell1,
		Cell2:   responderCell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, responderData); err != nil {
		return err
	}

	if err = SetRowHeight(title.file, title.sheet, responder, unitTitleHeight, title.nextRowID); err != nil {
		return err
	}

	// reviewers
	reviewersCell1, reviewersCell2, _ := Get2Cells(1, title.nextRowID+1, 6, title.nextRowID+1)

	var reviewersList string
	rowRevID := title.nextRowID + 1
	for _, reviewer := range reviewers {
		reviewersList = "- " + reviewer + "\n"
		if err = SetRowHeight(title.file, title.sheet, reviewersList, unitHeightForMO, rowRevID); err != nil {
			return err
		}
	}

	if rowRevID == 13+2*len(t.Responses) {
		style, err = setTitleResponsesStyle(title.file, 10, 2, false, false, true, true, true)
		if err != nil {
			return err
		}
	}

	reviewersData := Cell{
		Data:    reviewersList,
		Cell1:   reviewersCell1,
		Cell2:   reviewersCell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, reviewersData); err != nil {
		return err
	}
	return nil
}

func (title *title) printDeadlines(deadline string) error {
	cell1, cell2, _ := Get2Cells(7, title.nextRowID, 8, title.nextRowID+1)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    deadline,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	return nil
}

func (title *title) printFirstHeader() error {
	cell1, cell2, _ := Get2Cells(3, 1, 12, 1)

	style, err := setTitleStyle(title.file, 10, true, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    titleChapter,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	return nil
}

func (title *title) printSecondHeader() error {
	cell1, cell2, _ := Get2Cells(3, 3, 12, 3)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    titleFirstHeader,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	return nil
}

func (title *title) printThirdHeader() error {
	cell1, cell2, _ := Get2Cells(2, 5, 13, 5)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    titleSecondHeader,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}
	if err = SetRowHeight(title.file, title.sheet, titleSecondHeader, unitHeightForMO, 5); err != nil {
		return err
	}

	return nil
}

func (title *title) printFourthHeader() error {
	cell1, cell2, _ := Get2Cells(3, 7, 12, 7)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    titleThirdHeader,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}
	if err = SetRowHeight(title.file, title.sheet, titleThirdHeader, unitTitleHeight, 7); err != nil {
		return err
	}

	return nil
}

func (title *title) setColWidths() error {
	// ширина столбцов указана в мм
	m := make(map[string]float64)
	m["A"] = 12
	m["B"] = 10
	m["C"] = 8
	m["D"] = 16
	m["E"] = 62
	m["F"] = 37
	m["G"] = 42
	m["H"] = 14
	m["I"] = 10
	m["J"] = 10
	m["K"] = 10
	m["L"] = 20
	m["M"] = 10
	m["N"] = 14
	m["O"] = 10

	for k, v := range m {
		width := v / 2.5
		if err := title.file.SetColWidth(title.sheet, k, k, width); err != nil {
			return err
		}
	}

	return nil
}

func (title *title) printTitleFooterFirstHeader(region string) error {
	cell1, cell2, _ := Get2Cells(1, title.nextRowID, 15, title.nextRowID)

	style, err := setTitleHeaderFooterStyle(title.file, 10, true, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    footerFirstHeader + region,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}

	if err = SetRowHeight(title.file, title.sheet, cell.Data, unitTitleHeight, title.nextRowID); err != nil {
		return err
	}
	title.nextRowID++
	return nil
}

func (title *title) printTitleFooterSecondHeader() error {
	cell1, cell2, _ := Get2Cells(1, title.nextRowID, 15, title.nextRowID)

	style, err := setTitleHeaderFooterStyle(title.file, 10, true, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    footerSecondHeader,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}
	title.nextRowID++
	return nil
}

func (title *title) printTitleFooterThirdHeader() error {
	cell1, cell2, _ := Get2Cells(1, title.nextRowID, 3, title.nextRowID+1)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    footerThirdHeader,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}
	if err := SetRowHeight(title.file, title.sheet, footerThirdHeader, 15, title.nextRowID+1); err != nil {
		return err
	}
	return nil
}

func (title *title) printTitleFooterFourthHeader() error {
	cell1, cell2, _ := Get2Cells(4, title.nextRowID, 15, title.nextRowID)

	style, err := setTitleStyle(title.file, 10, false, true)
	if err != nil {
		return err
	}

	cell := Cell{
		Data:    footerFourthHeader,
		Cell1:   cell1,
		Cell2:   cell2,
		StyleID: style,
	}

	if err = printAnyCell(title.file, title.sheet, cell); err != nil {
		return err
	}
	title.nextRowID++
	return nil
}

func (title *title) printTitleFooterPreHeader() error {
	lCols := [4]int{4, 6, 8}
	rCols := [4]int{5, 7, 15}
	orgsToOKPO := [4]string{footerPreHeader, "", ""}

	for i := 0; i < 3; i++ {
		cell1, _ := excelize.CoordinatesToCellName(lCols[i], title.nextRowID)
		cell2, _ := excelize.CoordinatesToCellName(rCols[i], title.nextRowID)

		style, err := setTitleStyle(title.file, 10, false, true)
		if err != nil {
			return err
		}

		cell := Cell{
			Data:    orgsToOKPO[i],
			Cell1:   cell1,
			Cell2:   cell2,
			StyleID: style,
		}

		if err = printAnyCell(title.file, title.sheet, cell); err != nil {
			return err
		}
	}
	title.nextRowID++
	return nil
}

func (title *title) printTitleFooterLabel() error {
	lCols := [5]int{1, 4, 6, 8}
	rCols := [5]int{3, 5, 7, 15}
	footerLables := [5]string{"1", "2", "3", "4"}

	for i := 0; i < 4; i++ {
		cell1, cell2, _ := Get2Cells(lCols[i], title.nextRowID, rCols[i], title.nextRowID)

		style, err := setTitleStyle(title.file, 10, false, true)
		if err != nil {
			return err
		}

		cell := Cell{
			Data:    footerLables[i],
			Cell1:   cell1,
			Cell2:   cell2,
			StyleID: style,
		}

		if err = printAnyCell(title.file, title.sheet, cell); err != nil {
			return err
		}
	}
	title.nextRowID++
	return nil
}

func (title *title) printTitleFooterData(okud, okpo string) error {
	lCols := [5]int{1, 4, 6, 8}
	rCols := [5]int{3, 5, 7, 15}
	codes := [5]string{okud, okpo, "", ""}

	for i := 0; i < 4; i++ {
		cell1, cell2, _ := Get2Cells(lCols[i], title.nextRowID, rCols[i], title.nextRowID)

		style, err := setTitleStyle(title.file, 10, false, true)
		if err != nil {
			return err
		}

		cell := Cell{
			Data:    codes[i],
			Cell1:   cell1,
			Cell2:   cell2,
			StyleID: style,
		}

		if err = printAnyCell(title.file, title.sheet, cell); err != nil {
			return err
		}
	}

	return nil
}
