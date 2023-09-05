package internal

import "github.com/xuri/excelize/v2"

type nsi struct {
	Sheet   string
	ColID   int
	Address string
	Profile string
}

type dropList struct {
	initialValue   interface{}
	dataValidation *excelize.DataValidation
	isError        bool
}

type cellValue struct {
	value   interface{}
	isError bool
}
