package internal

import (
	"errors"
	"fmt"
	"github.com/jmoiron/sqlx"
	"github.com/xuri/excelize/v2"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/excel/form/models"
	"log"
	"strconv"
	"strings"
)

type ichapter struct {
	sheet     string
	nextRowID int
	nsi       nsi
	doc       DOC
	file      *excelize.File
	dvMap     map[int]*excelize.DataValidation
	Errors    [][]interface{}
}

func newChapter(doc DOC, f *Form, n NSI) *ichapter {
	return &ichapter{
		sheet:     doc.LABEL,
		nextRowID: 1,
		nsi: nsi{
			Sheet:   f.NSISheet,
			ColID:   f.NSIColID,
			Address: n.Address,
			Profile: n.Profile,
		},
		doc:    doc,
		file:   f.File,
		dvMap:  make(map[int]*excelize.DataValidation),
		Errors: nil,
	}
}

// PrintChapter печатает раздел
func PrintChapter(f *Form, doc DOC, data []string, nsi NSI, db *sqlx.DB) error {
	chptr := newChapter(doc, f, nsi)

	sheetOpts := SheetOptions{
		Protected: doc.COMMON.STATICROWS,
		HasErrors: doc.ERROR,
	}

	if err := CreateSheet(f.File, chptr.sheet, sheetOpts); err != nil {
		return err
	}

	if err := chptr.printCode(); err != nil {
		return err
	}

	if err := chptr.printChapterText(); err != nil {
		return err
	}

	if err := chptr.printPreHeader(); err != nil {
		return err
	}

	if err := chptr.printHeader(); err != nil {
		return err
	}

	if err := chptr.printColumns(); err != nil {
		return err
	}

	if len(data) != 0 {
		chapterID := chptr.doc.ID
		formulas, err := getStatFormulas(db, chapterID)
		if err != nil {
			return err
		}

		if existsRowNum(doc) {
			if err := chptr.printData(data, formulas, f.ChapterMap); err != nil {
				return err
			}
		} else {
			if err := chptr.printDataWithoutRowNum(data, formulas, f.ChapterMap); err != nil {
				return err
			}
		}

		if !doc.COMMON.STATICROWS {
			if err := chptr.printExtraDropLists(); err != nil {
				return err
			}
		}
	}

	if err := chptr.printFooterCaption(); err != nil {
		return err
	}

	if err := chptr.printFooter(); err != nil {
		return err
	}

	if err := SetColWidth(chptr.file, chptr.sheet, doc); err != nil {
		return err
	}

	f.Errors = append(f.Errors, chptr.Errors...)

	return nil
}

func (chapter *ichapter) printCode() error {
	if chapter.doc.CODE == "" {
		return nil
	}

	cell, err := excelize.CoordinatesToCellName(len(chapter.doc.COLUMNS), chapter.nextRowID)
	if err != nil {
		return err
	}

	style, err := setCommonStyle(&Style{}, chapter.doc, chapter.file)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellStyle(chapter.sheet, cell, cell, style); err != nil {
		return err
	}

	if err = chapter.file.SetCellValue(chapter.sheet, cell, chapter.doc.CODE); err != nil {
		return err
	}

	if err = SetRowHeight(chapter.file, chapter.sheet, chapter.doc.CODE,
		chapter.doc.COMMON.ROWUNITWIDTH, chapter.nextRowID); err != nil {
		return err
	}

	chapter.nextRowID += 1

	return nil
}

func (chapter *ichapter) printChapterText() error {

	if chapter.doc.CHAPTER.TEXT == "" {
		return nil
	}

	chapterStyle := Style{
		Alignment: Alignment{
			Horizontal: chapter.doc.CHAPTER.STYLE.HORALIGNMENT,
			Vertical:   chapter.doc.CHAPTER.STYLE.VERTALIGNMENT,
			Indent:     chapter.doc.CHAPTER.STYLE.INDENT,
		},
		Font: Font{
			Family: chapter.doc.CHAPTER.STYLE.FONT,
			Size:   chapter.doc.CHAPTER.STYLE.SIZE,
			Bold:   chapter.doc.CHAPTER.STYLE.FONT_STYLE.BOLD,
			Italic: chapter.doc.CHAPTER.STYLE.FONT_STYLE.ITALIC,
		},
	}
	cell1, err := excelize.CoordinatesToCellName(1, chapter.nextRowID)
	if err != nil {
		return err
	}

	cell2, err := excelize.CoordinatesToCellName(len(chapter.doc.COLUMNS), chapter.nextRowID)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellValue(chapter.sheet, cell1, chapter.doc.CHAPTER.TEXT); err != nil {
		return err
	}

	if err = chapter.file.MergeCell(chapter.sheet, cell1, cell2); err != nil {
		return err
	}

	style, err := setCommonStyle(&chapterStyle, chapter.doc, chapter.file)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellStyle(chapter.sheet, cell1, cell2, style); err != nil {
		return err
	}

	if err = SetRowHeight(chapter.file, chapter.sheet, chapter.doc.CHAPTER.TEXT,
		chapter.doc.COMMON.ROWUNITWIDTH, chapter.nextRowID); err != nil {
		return err
	}

	chapter.nextRowID += 1

	return nil

}
func (chapter *ichapter) printPreHeader() error {

	if chapter.doc.PREHEADER.TEXT == "" {
		return nil
	}

	cell1, err := excelize.CoordinatesToCellName(1, chapter.nextRowID)
	if err != nil {
		return err
	}

	cell2, err := excelize.CoordinatesToCellName(len(chapter.doc.COLUMNS), chapter.nextRowID)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellValue(chapter.sheet, cell1, chapter.doc.PREHEADER.TEXT); err != nil {
		return err
	}

	if err = chapter.file.MergeCell(chapter.sheet, cell1, cell2); err != nil {
		return err
	}

	preHeaderStyle := Style{
		Alignment: Alignment{
			Horizontal: chapter.doc.PREHEADER.STYLE.HORALIGNMENT,
			Vertical:   chapter.doc.PREHEADER.STYLE.VERTALIGNMENT,
			Indent:     chapter.doc.PREHEADER.STYLE.INDENT,
		},
		Font: Font{
			Size:   chapter.doc.PREHEADER.STYLE.SIZE,
			Family: chapter.doc.PREHEADER.STYLE.FONT,
			Bold:   chapter.doc.PREHEADER.STYLE.FONT_STYLE.BOLD,
			Italic: chapter.doc.PREHEADER.STYLE.FONT_STYLE.ITALIC,
		},
	}

	style, err := setCommonStyle(&preHeaderStyle, chapter.doc, chapter.file)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellStyle(chapter.sheet, cell1, cell2, style); err != nil {
		return err
	}

	if err = SetRowHeight(chapter.file, chapter.sheet, chapter.doc.PREHEADER.TEXT,
		chapter.doc.COMMON.ROWUNITWIDTH, chapter.nextRowID); err != nil {
		return err
	}

	chapter.nextRowID += 1

	return nil
}

func (chapter *ichapter) printHeader() error {

	if len(chapter.doc.HEADER.HEADERS) == 0 {
		return nil
	}

	headerStyle := Style{
		Alignment: Alignment{
			Horizontal: chapter.doc.HEADER.STYLE.HORALIGNMENT,
			Vertical:   chapter.doc.HEADER.STYLE.VERTALIGNMENT,
			Indent:     chapter.doc.HEADER.STYLE.INDENT,
		},
		Font: Font{
			Family: chapter.doc.HEADER.STYLE.FONT,
			Size:   chapter.doc.HEADER.STYLE.SIZE,
			Bold:   chapter.doc.HEADER.STYLE.FONT_STYLE.BOLD,
			Italic: chapter.doc.HEADER.STYLE.FONT_STYLE.ITALIC,
		},
	}

	style, err := setCommonStyle(&headerStyle, chapter.doc, chapter.file)
	if err != nil {
		return err
	}

	maxNewLines := make([]int, chapter.doc.HEADER.HEADERROWS)

	for _, header := range chapter.doc.HEADER.HEADERS {

		cell1, err := excelize.CoordinatesToCellName(header.HEADERGRID[1], chapter.nextRowID+header.HEADERGRID[0]-1)
		if err != nil {
			return err
		}

		cell2, err := excelize.CoordinatesToCellName(header.HEADERGRID[3], chapter.nextRowID+header.HEADERGRID[2]-1)
		if err != nil {
			return err
		}

		if err = chapter.file.SetCellValue(chapter.sheet, cell1, header.HEADERCAPTION); err != nil {
			return err
		}

		if err = chapter.file.MergeCell(chapter.sheet, cell1, cell2); err != nil {
			return err
		}

		if err = chapter.file.SetCellStyle(chapter.sheet, cell1, cell2, style); err != nil {
			return err
		}

		curNewLines := CountNewLines(header.HEADERCAPTION)
		if curNewLines > maxNewLines[header.HEADERGRID[0]-1] {
			maxNewLines[header.HEADERGRID[0]-1] = curNewLines
			maxNewLines[header.HEADERGRID[2]-1] = curNewLines
		}
	}

	unitHeight := chapter.doc.COMMON.ROWUNITWIDTH
	for i, n := range maxNewLines {
		height := float64(n * unitHeight)
		if err = chapter.file.SetRowHeight(chapter.sheet, chapter.nextRowID+i, height); err != nil {
			return err
		}
	}

	chapter.nextRowID += chapter.doc.HEADER.HEADERROWS

	return nil
}

func (chapter *ichapter) printColumns() error {

	if len(chapter.doc.COLUMNS) == 0 {
		return nil
	}

	colStyle := Style{
		Alignment: Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Font: Font{
			Bold:   chapter.doc.HEADER.STYLE.FONT_STYLE.BOLD,
			Italic: chapter.doc.HEADER.STYLE.FONT_STYLE.ITALIC,
		},
	}

	style, err := setCommonStyle(&colStyle, chapter.doc, chapter.file)
	if err != nil {
		return err
	}

	for i, column := range chapter.doc.COLUMNS {

		cell, err := excelize.CoordinatesToCellName(i+1, chapter.nextRowID)
		if err != nil {
			return err
		}

		if err = chapter.file.SetCellStyle(chapter.sheet, cell, cell, style); err != nil {
			return err
		}

		if err = chapter.file.SetCellValue(chapter.sheet, cell, column.COLUMNNUM); err != nil {
			return err
		}

		unitHeight := chapter.doc.COMMON.ROWUNITWIDTH
		if err = SetRowHeight(chapter.file, chapter.sheet, column.COLUMNNUM, unitHeight, chapter.nextRowID); err != nil {
			return err
		}

	}

	chapter.nextRowID += 1

	return nil
}

func (chapter *ichapter) printData(data []string, formulas []models.Formula, chapterMap map[int]ChapterDetails) error {
	// map[row_num]row_id
	dataRowMap := make(map[int]int)
	idx := 0

	if len(chapter.doc.ROWBREAKS) == 0 {
		log.Println("rowBreaks is empty")
		rowID := chapter.nextRowID
		for i := range chapter.doc.DATA {
			cellsData, err := chapter.getColTypes(data[i], chapter.dvMap)
			if err != nil {
				return err
			}
			if err = chapter.printCellData(cellsData, rowID); err != nil {
				return err
			}
			dataRowMap[i+1] = rowID
			rowID++
		}
	} else {
		if chapter.doc.ROWBREAKS[idx].PREVROW == 0 {
			if err := chapter.printRowBreak(chapter.nextRowID, idx); err != nil {
				return err
			}
			chapter.nextRowID, idx = chapter.nextRowID+1, idx+1
		}
		rowID := chapter.nextRowID
		for i := range chapter.doc.DATA {
			//rowID := nextRowID + i
			if idx == len(chapter.doc.ROWBREAKS) {
				idx = idx - 1
			}

			rowNum, err := getRowNumVal(chapter.doc.DATA[i]["ROW_NUM"])
			if err != nil {
				return err
			}

			if rowNum == chapter.doc.ROWBREAKS[idx].PREVROW+1 && chapter.doc.ROWBREAKS[idx].PREVROW != 0 {
				if err := chapter.printRowBreak(rowID, idx); err != nil {
					return err
				}
				rowID, idx = rowID+1, idx+1
			}
			cellsData, err := chapter.getColTypes(data[i], chapter.dvMap)
			if err != nil {
				return err
			}

			if err = chapter.printCellData(cellsData, rowID); err != nil {
				return err
			}
			dataRowMap[rowNum] = rowID
			rowID++
		}
	}

	templateID, sheetName := chapterMap[chapter.doc.ID].TemplateID, chapterMap[chapter.doc.ID].SheetName

	chapterMap[chapter.doc.ID] = ChapterDetails{
		TemplateID: templateID,
		SheetName:  sheetName,
		DataRowMap: dataRowMap,
	}

	if err := printFormula(formulas, dataRowMap, chapterMap, chapter); err != nil {
		return err
	}

	if err := setCellStyles(chapter.file, chapter.sheet, chapter.doc, dataRowMap); err != nil {
		return err
	}

	if err := setRowUnion(chapter.file, chapter.sheet, chapter.doc, dataRowMap); err != nil {
		return err
	}

	chapter.nextRowID += len(chapter.doc.DATA) + len(chapter.doc.ROWBREAKS)

	return nil
}

func printFormula(formulas []models.Formula, dataRowMap map[int]int, chapterMap map[int]ChapterDetails, chapter *ichapter) error {
	for _, f := range formulas {
		if f.TargetCells[0] != -1 {
			target, err := excelize.CoordinatesToCellName(f.TargetCells[1], dataRowMap[f.TargetCells[0]])
			if err != nil {
				format := `invalid formula: chapter_template_id=%d, target=%v, source=%v`
				log.Printf(format, chapterMap[chapter.doc.ID].TemplateID,
					f.TargetCells, f.SourceCells)
				continue
			}

			layout := f.Layout

			for _, src := range f.SourceCells {
				var source string
				if len(src) == 1 {
					source = strconv.Itoa(src[0])
				} else if len(src) == 2 {
					source, err = excelize.CoordinatesToCellName(src[1], dataRowMap[src[0]])
					if err != nil {
						format := `invalid formula: chapter_template_id=%d, target=%v, source=%v`
						log.Printf(format, chapterMap[chapter.doc.ID].TemplateID,
							f.TargetCells, f.SourceCells)
						break
					}
				} else if len(src) == 3 {
					var (
						sheetName string
						rowMap    = make(map[int]int)
					)

					templateID := src[2]

					for _, detail := range chapterMap {
						if detail.TemplateID == templateID {
							sheetName = detail.SheetName
							if sheetName == "" {
								return fmt.Errorf("sheet name is nil")
							}

							for k, v := range detail.DataRowMap {
								rowMap[k] = v
							}
						}
					}

					cell, err := excelize.CoordinatesToCellName(src[1], rowMap[src[0]])
					if err != nil {
						format := `invalid formula: chapter_template_id=%d, target=%v, source=%v`
						log.Printf(format, chapterMap[chapter.doc.ID].TemplateID,
							f.TargetCells, f.SourceCells)
						break
					}

					source = fmt.Sprintf("'%s'!%s", sheetName, cell)
				} else {
					errMsg := fmt.Sprintf("invalid source: %v, please check source is valid!", src)
					return errors.New(errMsg)

				}

				layout = strings.Replace(layout, "%s", source, 1)
			}
			if err = chapter.file.SetCellFormula(chapter.sheet, target, layout); err != nil {
				return err
			}

			chapter.doc.STYLES = append(chapter.doc.STYLES, ExtraStyle{
				GRID: []int{
					f.TargetCells[0], f.TargetCells[1],
					f.TargetCells[0], f.TargetCells[1],
				},
				PROTECTED: true,
			})

		} else {
			for i := 0; i < len(chapter.doc.DATA); i++ {

				target, err := excelize.CoordinatesToCellName(f.TargetCells[1], chapter.nextRowID+i)
				if err != nil {
					format := `invalid formula: chapter_template_id=%d, target=%v, source=%v`
					log.Printf(format, chapterMap[chapter.doc.ID].TemplateID,
						f.TargetCells, f.SourceCells)
					continue
				}

				layout := f.Layout

				for _, src := range f.SourceCells {
					var source string
					if len(src) == 1 {
						source = strconv.Itoa(src[0])
					} else if len(src) == 2 {
						if src[0] != -1 {
							source, err = excelize.CoordinatesToCellName(src[1], dataRowMap[src[0]])
							if err != nil {
								format := `invalid formula: chapter_template_id=%d, target=%v, source=%v`
								log.Printf(format, chapterMap[chapter.doc.ID].TemplateID,
									f.TargetCells, f.SourceCells)
								break
							}
						} else {
							source, err = excelize.CoordinatesToCellName(src[1], chapter.nextRowID+i)
							if err != nil {
								format := `invalid formula: chapter_template_id=%d, target=%v, source=%v`
								log.Printf(format, chapterMap[chapter.doc.ID].TemplateID,
									f.TargetCells, f.SourceCells)
								break
							}
						}

					} else if len(src) == 3 {
						var (
							sheetName string
							rowMap    = make(map[int]int)
						)

						templateID := src[2]

						for _, detail := range chapterMap {
							if detail.TemplateID == templateID {
								sheetName = detail.SheetName
								if sheetName == "" {
									return fmt.Errorf("sheet name is nil")
								}

								for k, v := range detail.DataRowMap {
									rowMap[k] = v
								}
							}
						}

						cell, err := excelize.CoordinatesToCellName(src[1], rowMap[src[0]])
						if err != nil {
							format := `invalid formula: chapter_template_id=%d, target=%v, source=%v`
							log.Printf(format, chapterMap[chapter.doc.ID].TemplateID,
								f.TargetCells, f.SourceCells)
							break
						}

						source = fmt.Sprintf("'%s'!%s", sheetName, cell)
					} else {
						errMsg := fmt.Sprintf("invalid source: %v, please check source is valid!", src)
						return errors.New(errMsg)

					}

					layout = strings.Replace(layout, "%s", source, 1)
				}
				if err = chapter.file.SetCellFormula(chapter.sheet, target, layout); err != nil {
					return err
				}

				chapter.doc.STYLES = append(chapter.doc.STYLES, ExtraStyle{
					GRID: []int{
						i + 1, f.TargetCells[1],
						i + 1, f.TargetCells[1],
					},
					PROTECTED: true,
				})
			}
		}

	}

	return nil
}

func (chapter *ichapter) printDataWithoutRowNum(data []string, formulas []models.Formula, chapterMap map[int]ChapterDetails) error {
	rowID := chapter.nextRowID
	for i := range data {
		cellsData, err := chapter.getColTypes(data[i], chapter.dvMap)
		if err != nil {
			return err
		}
		if err = chapter.printCellData(cellsData, rowID); err != nil {
			return err
		}
		rowID++
	}

	if err := printFormulaWithoutRowNum(formulas, chapterMap, chapter); err != nil {
		return err
	}

	chapter.nextRowID += len(data)

	return nil
}

func printFormulaWithoutRowNum(formulas []models.Formula, chapterMap map[int]ChapterDetails, chapter *ichapter) error {
	for _, f := range formulas {
		if f.TargetCells[0] != -1 {
			target, err := excelize.CoordinatesToCellName(f.TargetCells[1], f.TargetCells[0]+chapter.nextRowID-1)
			if err != nil {
				return err
			}

			layout := f.Layout

			for _, src := range f.SourceCells {
				var source string
				switch len(src) {
				case 1:
					source = strconv.Itoa(src[0])
				case 2:
					source, err = excelize.CoordinatesToCellName(src[1], src[0]+chapter.nextRowID-1)
					if err != nil {
						return err
					}
				case 3:
					cell, err := excelize.CoordinatesToCellName(src[1], src[0]+chapter.nextRowID-1)
					if err != nil {
						return err
					}
					var (
						sheetName string
					)

					templateID := src[2]

					for _, detail := range chapterMap {
						if detail.TemplateID == templateID {
							sheetName = detail.SheetName
							if sheetName == "" {
								return fmt.Errorf("sheet name is nil")
							}
							break
						}
					}
					//sheetName = chapterMap[src[2]].SheetName
					source = fmt.Sprintf("'%s'!%s", sheetName, cell)
				default:
					errMsg := fmt.Sprintf("invalid source: %v, please check source is valid!", src)
					return errors.New(errMsg)
				}

				layout = strings.Replace(layout, "%s", source, 1)
			}

			if err = chapter.file.SetCellFormula(chapter.sheet, target, layout); err != nil {
				return err
			}

			chapter.doc.STYLES = append(chapter.doc.STYLES, ExtraStyle{
				GRID: []int{
					f.TargetCells[0], f.TargetCells[1],
					f.TargetCells[0], f.TargetCells[1],
				},
				PROTECTED: true,
			})
		} else {
			for i := 0; i < len(chapter.doc.DATA); i++ {
				target, err := excelize.CoordinatesToCellName(f.TargetCells[1], chapter.nextRowID+i)
				if err != nil {
					return err
				}

				layout := f.Layout

				for _, src := range f.SourceCells {
					var source string
					switch len(src) {
					case 1:
						source = strconv.Itoa(src[0])
					case 2:
						if src[0] == -1 {
							source, err = excelize.CoordinatesToCellName(src[1], chapter.nextRowID+i)
							if err != nil {
								return err
							}
						} else {
							source, err = excelize.CoordinatesToCellName(src[1], src[0]+chapter.nextRowID-1)
							if err != nil {
								return err
							}
						}
					case 3:
						cell, err := excelize.CoordinatesToCellName(src[1], src[0]+chapter.nextRowID-1)
						if err != nil {
							return err
						}
						var (
							sheetName string
						)

						templateID := src[2]

						for _, detail := range chapterMap {
							if detail.TemplateID == templateID {
								sheetName = detail.SheetName
								if sheetName == "" {
									return fmt.Errorf("sheet name is nil")
								}
								break
							}
						}
						//sheetName = chapterMap[src[2]].SheetName
						source = fmt.Sprintf("'%s'!%s", sheetName, cell)
					default:
						errMsg := fmt.Sprintf("invalid source: %v, please check source is valid!", src)
						return errors.New(errMsg)
					}

					layout = strings.Replace(layout, "%s", source, 1)
				}

				if err = chapter.file.SetCellFormula(chapter.sheet, target, layout); err != nil {
					return err
				}

				chapter.doc.STYLES = append(chapter.doc.STYLES, ExtraStyle{
					GRID: []int{
						i + 1, f.TargetCells[1],
						i + 1, f.TargetCells[1],
					},
					PROTECTED: true,
				})
			}
		}
	}

	return nil
}

func (chapter *ichapter) printCellData(cellsData []interface{}, rowID int) error {
	maxNewLines := 1
	for j, cellData := range cellsData {
		if cellData == nil {
			continue
		}

		cell, err := excelize.CoordinatesToCellName(j+1, rowID)
		if err != nil {
			return err
		}

		var (
			value interface{}
			color string
		)

		switch v := cellData.(type) {
		case *dropList:
			value = v.initialValue
			if v.isError {
				color = "#FF0000"
				chapter.Errors = append(chapter.Errors, []interface{}{chapter.sheet, cell, value})
			}
			dv := &excelize.DataValidation{
				Sqref:    cell + ":" + cell,
				Formula1: v.dataValidation.Formula1,
				Type:     v.dataValidation.Type,
			}
			if err = chapter.file.AddDataValidation(chapter.sheet, dv); err != nil {
				return err
			}
		case *cellValue:
			value = v.value
			if v.isError {
				color = "#FF0000"
				chapter.Errors = append(chapter.Errors, []interface{}{chapter.sheet, cell, value})
			}
		}

		if err = chapter.file.SetCellValue(chapter.sheet, cell, value); err != nil {
			return err
		}
		curNewLines := CountNewLines(value)
		if curNewLines > maxNewLines {
			maxNewLines = curNewLines
		}

		dataStyle := Style{
			Alignment: Alignment{
				Horizontal: chapter.doc.COLUMNS[j].STYLE.HORALIGNMENT,
				Vertical:   chapter.doc.COLUMNS[j].STYLE.VERTALIGNMENT,
				Indent:     chapter.doc.COLUMNS[j].STYLE.INDENT,
			},
			Font: Font{
				Size:   chapter.doc.COLUMNS[j].STYLE.SIZE,
				Family: chapter.doc.COLUMNS[j].STYLE.FONT,
				Bold:   chapter.doc.COLUMNS[j].STYLE.FONT_STYLE.BOLD,
				Italic: chapter.doc.COLUMNS[j].STYLE.FONT_STYLE.ITALIC,
			},
			Fill: Fill{
				Color: color,
			},
		}

		style, err := setCommonStyle(&dataStyle, chapter.doc, chapter.file)
		if err != nil {
			return err
		}

		if err = chapter.file.SetCellStyle(chapter.sheet, cell, cell, style); err != nil {
			return err
		}
	}

	unitHeight := chapter.doc.COMMON.ROWUNITWIDTH
	height := float64(maxNewLines * unitHeight)
	if err := chapter.file.SetRowHeight(chapter.sheet, rowID, height); err != nil {
		return err
	}

	return nil
}

func (chapter *ichapter) printExtraDropLists() error {
	firstRowID := chapter.nextRowID

	for colID, dv := range chapter.dvMap {
		cell1, cell2, err := Get2Cells(colID, firstRowID, colID, firstRowID+50)
		if err != nil {
			return err
		}
		dv.Sqref = cell1 + ":" + cell2
		if err = chapter.file.AddDataValidation(chapter.sheet, dv); err != nil {
			return err
		}
	}

	return nil
}

func (chapter *ichapter) printRowBreak(rowID, idx int) error {
	cell1, err := excelize.CoordinatesToCellName(1, rowID)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellValue(chapter.sheet, cell1, chapter.doc.ROWBREAKS[idx].BREAKCAPTION); err != nil {
		return err
	}
	rowBreakStyle := Style{
		Alignment: Alignment{
			Horizontal: chapter.doc.ROWBREAKS[idx].STYLE.HORALIGNMENT,
			Vertical:   chapter.doc.ROWBREAKS[idx].STYLE.VERTALIGNMENT,
			Indent:     chapter.doc.ROWBREAKS[idx].INDENT,
		},
		Font: Font{
			Size:   chapter.doc.ROWBREAKS[idx].STYLE.SIZE,
			Family: chapter.doc.ROWBREAKS[idx].STYLE.FONT,
			Bold:   chapter.doc.ROWBREAKS[idx].STYLE.FONT_STYLE.BOLD,
			Italic: chapter.doc.ROWBREAKS[idx].STYLE.FONT_STYLE.ITALIC,
		},
	}

	style, err := setCommonStyle(&rowBreakStyle, chapter.doc, chapter.file)
	if err != nil {
		return err
	}

	cell2, err := excelize.CoordinatesToCellName(len(chapter.doc.COLUMNS), rowID)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellStyle(chapter.sheet, cell1, cell2, style); err != nil {
		return err
	}

	if err = SetRowHeight(chapter.file,
		chapter.sheet, chapter.doc.ROWBREAKS[idx].BREAKCAPTION,
		chapter.doc.COMMON.ROWUNITWIDTH, rowID); err != nil {
		return err
	}

	return nil
}

func (chapter *ichapter) printFooterCaption() error {
	cell, err := excelize.CoordinatesToCellName(1, chapter.nextRowID)
	if err != nil {
		return err
	}
	if err = chapter.file.SetCellValue(chapter.sheet, cell, chapter.doc.FOOTER.CAPTION); err != nil {
		return err
	}

	if err = SetRowHeight(chapter.file,
		chapter.sheet, chapter.doc.FOOTER.CAPTION,
		chapter.doc.COMMON.ROWUNITWIDTH, chapter.nextRowID); err != nil {
		return err
	}

	return err
}

func (chapter *ichapter) printFooter() error {

	if len(chapter.doc.FOOTER.DATA) == 0 {
		return nil
	}

	paramCell1, err := excelize.CoordinatesToCellName(1, chapter.nextRowID)
	if err != nil {
		return err
	}

	for i := 0; i < len(chapter.doc.FOOTER.DATA); i++ {

		unitHeight := chapter.doc.COMMON.ROWUNITWIDTH
		if err = SetRowHeight(chapter.file, chapter.sheet, chapter.doc.FOOTER.DATA[i].VALUE, unitHeight, chapter.nextRowID+i); err != nil {
			return err
		}

		valueCell, err := excelize.CoordinatesToCellName(ToInt(chapter.doc.FOOTER.DATA[i].COLUMNNUM), chapter.nextRowID+chapter.doc.FOOTER.DATA[i].ROWNUM)
		if err != nil {
			return err
		}

		if err = chapter.file.SetCellValue(chapter.sheet, valueCell, chapter.doc.FOOTER.DATA[i].VALUE); err != nil {
			return err
		}
	}

	cell2, err := excelize.CoordinatesToCellName(ToInt(chapter.doc.FOOTER.DATA[len(chapter.doc.FOOTER.DATA)-1].COLUMNNUM), chapter.nextRowID+chapter.doc.FOOTER.DATA[len(chapter.doc.FOOTER.DATA)-1].ROWNUM)
	if err != nil {
		return err
	}

	styleID, err := setFooterStyle(&Style{}, chapter.doc, chapter.file)
	if err != nil {
		return err
	}

	if err = chapter.file.SetCellStyle(chapter.sheet, paramCell1, cell2, styleID); err != nil {
		return err
	}

	return nil
}
