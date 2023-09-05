package internal

import (
	"bytes"
	"github.com/jmoiron/sqlx"
	"github.com/xuri/excelize/v2"
	"log"
)

type ChapterDetails struct {
	TemplateID int
	SheetName  string
	DataRowMap map[int]int
}

type Form struct {
	ErrorsSheet string
	ErrorsRowID int
	Errors      [][]interface{}
	NSISheet    string
	NSIColID    int
	ChapterMap  map[int]ChapterDetails
	File        *excelize.File
}

func newForm(file *excelize.File) *Form {
	return &Form{
		ErrorsSheet: "Ошибки_данных",
		ErrorsRowID: 2,
		Errors:      nil,
		NSISheet:    "Справочная_информация",
		NSIColID:    1,
		ChapterMap:  make(map[int]ChapterDetails),
		File:        file,
	}
}

// PrintForm печатает форму отчёта
func PrintForm(request GetFormRequest, body []byte, nsi NSI, db *sqlx.DB) ([]byte, error) {
	// файл отчётности
	f := excelize.NewFile()

	// форма
	form := newForm(f)

	// титульный лист
	t := request.Title
	if t != nil {
		if err := PrintTitle(form, t); err != nil {
			return nil, err
		}
	} else {
		log.Println("Warning: empty title!")
	}

	if len(request.DOCS) == 0 {
		log.Println("Warning: empty docs")
	} else {
		if err := CreateDropListSheet(f, form.NSISheet); err != nil {
			return nil, err
		}

		for _, doc := range request.DOCS {
			templateID, err := getChapterTemplateID(db, doc.ID)
			if err != nil {
				return nil, err
			}
			form.ChapterMap[doc.ID] = ChapterDetails{
				TemplateID: templateID,
				SheetName:  doc.LABEL,
			}
		}
		// print chapters
		for i, doc := range request.DOCS {
			data := GetDataFromJSON(i, body)
			if err := PrintChapter(form, doc, data, nsi, db); err != nil {
				return nil, err
			}
		}
		// remove sheet created by default
		if err := f.DeleteSheet("Sheet1"); err != nil {
			return nil, err
		}

		if err := f.SetSheetVisible(form.NSISheet, false); err != nil {
			return nil, err
		}
	}

	if form.Errors != nil {
		if err := form.printErrors(); err != nil {
			return nil, err
		}
	}

	//write data from File to buffer
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		return nil, err
	}

	if err := f.Close(); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}
