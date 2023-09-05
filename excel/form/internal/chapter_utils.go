package internal

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"github.com/jmoiron/sqlx"
	"github.com/tidwall/gjson"
	"github.com/xuri/excelize/v2"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/excel/form/models"
	"gitlab.geogracom.com/skdf/skdf-excel-server-go/excel/form/pkg/parser"
	"io"
	"net/http"
	"strings"
)

func getStatFormulas(db *sqlx.DB, chapterID int) ([]models.Formula, error) {
	query := `select * from test_abac.get_stat_formula($1);`
	rows, err := db.Query(query, chapterID)
	if err != nil {
		return nil, err
	}

	var fs []models.Formula
	for rows.Next() {
		var target, expression string
		if err = rows.Scan(&target, &expression); err != nil {
			return nil, err
		}

		f, err := parser.ParseFormula(target, expression)
		if err != nil {
			return nil, err
		}

		fs = append(fs, models.Formula{
			TargetCells: f.Target,
			SourceCells: f.Source,
			Layout:      f.Layout,
		})
	}

	return fs, nil
}

func getChapterTemplateID(db *sqlx.DB, chapterID int) (int, error) {
	query := `select test_abac.get_chapter_template_id($1);`

	row := db.QueryRow(query, chapterID)
	if row.Err() != nil {
		return -1, row.Err()
	}

	var chapterTemplateID int
	if err := row.Scan(&chapterTemplateID); err != nil {
		return -1, err
	}

	if chapterTemplateID == 0 {
		return -1, fmt.Errorf("invalid chapter template id")
	}

	return chapterTemplateID, nil
}

func setRowUnion(f *excelize.File, sheet string, doc DOC, dataRowMap map[int]int) error {
	for _, union := range doc.ROWUNIONS {
		cell1, cell2, err := Get2Cells(
			union.GRID[1], dataRowMap[union.GRID[0]],
			union.GRID[3], dataRowMap[union.GRID[2]],
		)
		if err != nil {
			return err
		}

		if err = f.MergeCell(sheet, cell1, cell2); err != nil {
			return err
		}
	}
	return nil
}

func existsRowNum(doc DOC) bool {
	for _, col := range doc.COLUMNS {
		if col.VALUELABEL == "ROW_NUM" || col.VALUEFIELD == "ROW_NUM" {
			return true
		}
	}
	return false
}

func getNSIBody(address string, profile string, object string) ([]byte, error) {

	req, err := http.NewRequest(http.MethodGet, address+object, nil)
	if err != nil {
		return nil, err
	}

	req.Header.Add("Accept-Profile", profile)

	res, err := http.DefaultClient.Do(req)
	if err != nil {
		return nil, err
	}

	body, err := getBody(res.Body)
	if err != nil {
		return nil, err
	}

	return body, nil
}

func getBody(r io.ReadCloser) ([]byte, error) {
	var buf bytes.Buffer

	n, err := io.Copy(&buf, r)
	if err != nil {
		return nil, err
	}

	if err = r.Close(); err != nil {
		return nil, err
	}

	return buf.Bytes()[:n], nil
}

func getNSIValues(body []byte, label string) ([]string, error) {
	var nsiValues []map[string]interface{}
	if err := json.Unmarshal(body, &nsiValues); err != nil {
		return nil, err
	}

	names := make([]string, len(nsiValues))

	for i, nsiValue := range nsiValues {
		names[i] = nsiValue[label].(string)
	}

	return names, nil
}

func (chapter *ichapter) setSource(values []string) error {
	for i := 1; i <= len(values); i++ {
		cell, err := excelize.CoordinatesToCellName(chapter.nsi.ColID, i, true)
		if err != nil {
			return err
		}

		if err = chapter.file.SetCellValue(chapter.nsi.Sheet, cell, values[i-1]); err != nil {
			return err
		}
	}
	return nil
}

func (chapter *ichapter) getSourceRef(srcLen int) (string, error) {
	cell1, err := excelize.CoordinatesToCellName(chapter.nsi.ColID, 1, true)
	if err != nil {
		return "", err
	}
	cell2, err := excelize.CoordinatesToCellName(chapter.nsi.ColID, srcLen, true)
	if err != nil {
		return "", err
	}
	srcRef := fmt.Sprintf("%s!%s:%s", chapter.nsi.Sheet, cell1, cell2)
	chapter.nsi.ColID++
	return srcRef, nil
}

func setDataValidation(srcRef string) *excelize.DataValidation {
	dv := excelize.NewDataValidation(false)
	dv.SetSqrefDropList(srcRef)
	return dv
}

func (chapter *ichapter) setDropList(dataSrc, label string) (*dropList, error) {
	body, err := getNSIBody(chapter.nsi.Address, chapter.nsi.Profile, dataSrc)
	if err != nil {
		return nil, err
	}

	if label == "" {
		label = "name"
	}

	values, err := getNSIValues(body, strings.ToLower(label))
	if err != nil {
		return nil, err
	}

	if len(values) == 0 {
		return nil, errors.New("values not found for drop down list")
	}

	if err = chapter.setSource(values); err != nil {
		return nil, err
	}

	srcRef, err := chapter.getSourceRef(len(values))
	if err != nil {
		return nil, err
	}

	dv := setDataValidation(srcRef)

	return &dropList{
		initialValue:   "",
		dataValidation: dv,
	}, nil

}

func (chapter *ichapter) getColTypes(src string, dataValidationMap map[int]*excelize.DataValidation) ([]interface{}, error) {
	rows := make([]interface{}, len(chapter.doc.COLUMNS))
	for _, column := range chapter.doc.COLUMNS {
		var (
			dl  *dropList
			err error
		)

		if ((column.ISEDITABLE != nil && *column.ISEDITABLE) || (column.ISEDITABLE == nil)) &&
			(column.SELECTOR != nil && !column.ALLOWMULTI) &&
			(column.SELECTOR.DATASOURCE != "SUPPLIER") &&
			(column.SELECTOR.DATASOURCE != "CONTRACTOR") && column.SELECTOR.DATASOURCE != "" {
			if dataValidationMap[column.COLUMNID] == nil {
				dl, err = chapter.setDropList(column.SELECTOR.DATASOURCE, column.SELECTOR.VALUELABEL)
				if err != nil {
					return nil, err
				}
				dataValidationMap[column.COLUMNID] = dl.dataValidation
			} else {
				dl = &dropList{
					dataValidation: dataValidationMap[column.COLUMNID],
				}
			}
		}

		var path string
		if column.VALUELABEL != "" {
			path = getPath(column.VALUELABEL)
		} else {
			path = getPath(column.VALUEFIELD)
		}

		value, isError := parseValue(src, path, column.SEPARATOR)

		if dl != nil {
			rows[column.COLUMNID-1] = &dropList{
				initialValue:   value,
				dataValidation: dl.dataValidation,
				isError:        isError,
			}
		} else {
			rows[column.COLUMNID-1] = &cellValue{
				value:   value,
				isError: isError,
			}
		}
	}
	return rows, nil
}

func getPath(field string) string {
	strEscaper := strings.NewReplacer(
		":", "#",
		"[", ".",
		"]", "",
	)
	path := strEscaper.Replace(field)

	return path
}

func parseValue(src, path, separator string) (interface{}, bool) {
	var (
		value   interface{}
		isError bool
	)

	errPath := path + "_ERROR"
	errValue := gjson.Get(src, errPath).Value()
	if errValue != nil && !isEmptyArr(gjson.Get(src, errPath).Raw) {
		value = errValue
		isError = true
	} else {
		value = gjson.Get(src, path).Value()
	}

	switch obj := value.(type) {
	case int:
		return obj, isError
	case float64:
		return obj, isError
	case []interface{}:
		elems := make([]string, len(obj))
		if separator == "" {
			separator = ","
		}
		for i, item := range obj {
			if item != nil {
				elems[i] = fmt.Sprintf("%v", item)
			}
		}
		result := strings.Join(elems, separator)
		return result, isError
	default:
		return value, isError
	}

}

func isEmptyArr(raw string) bool {
	var res string
	escapeArr := strings.NewReplacer("[", "", "]", "")
	res = escapeArr.Replace(raw)
	return res == ""
}

func getRowNumVal(v interface{}) (int, error) {
	switch x := v.(type) {
	case float64:
		return int(x), nil
	case int:
		return x, nil
	}
	return -1, errors.New("invalid type of ROW_NUM")
}
