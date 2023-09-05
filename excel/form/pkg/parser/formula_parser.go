package parser

import (
	"errors"
	"fmt"
	"strconv"
	"strings"
)

type Formula struct {
	// Target example: [1][4] => {1, 4}
	Target []int

	// Source
	// example 1: [1][4] => {1, 4}
	// example 2: 7000 => {7000}, just number
	// example 3: [1][4][1233145] => {1, 4, 1233145},
	// where 1233145 is chapter_template_id
	// example 4: [#][1] => {-1, 1},
	// where (# = -1) means current row id
	Source [][]int

	// Layout example: [1][4]+[1][6] => "=%s+%s"
	Layout string
}

func ParseFormula(target, expression string) (*Formula, error) {
	targetCells, err := parseTarget(target)
	if err != nil {
		return nil, err
	}

	sourceCells, layout, err := parseExpression(expression)
	if err != nil {
		return nil, err
	}

	return &Formula{
		Target: targetCells,
		Source: sourceCells,
		Layout: layout,
	}, nil
}

func parseTarget(s string) ([]int, error) {
	indexes, err := parseIndexes(s)
	if err != nil {
		return nil, err
	}

	if len(indexes) != 2 {
		errMsg := fmt.Sprintf("invalid target: %s, please check target cell is valid!", s)
		return nil, errors.New(errMsg)
	}

	return indexes, nil
}

func parseExpression(s string) ([][]int, string, error) {
	var arr [][]int

	operands := getOperands(s)

	layout := s

	for _, operand := range operands {
		if !strings.Contains(operand, "[") && !strings.Contains(operand, "]") {
			num, err := strconv.Atoi(operand)
			if err != nil {
				return nil, "", err
			}
			arr = append(arr, []int{num})
		} else {
			idx, err := parseIndexes(operand)
			if err != nil {
				return nil, "", err
			}
			arr = append(arr, idx)
		}
		//log.Printf("layout: %s, operand: %s", layout, operand)
		layout = strings.Replace(layout, operand, "%s", 1)
	}

	if strings.Count(layout, "%s") != len(arr) {
		//log.Println(layout, arr)
		return nil, "", errors.New("formula parser is not working properly")
	}

	return arr, layout, nil
}

func parseIndexes(s string) ([]int, error) {
	var (
		arr      []int
		beg, end int
		opened   bool
	)

	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '[':
			if !opened {
				opened = true
				beg, end = i, -1
			} else {
				errMsg := fmt.Sprintf("invalid expression: %s, please check expression is valid!", s)
				return nil, errors.New(errMsg)
			}
		case ']':
			if opened {
				opened = false
				end = i
			} else {
				errMsg := fmt.Sprintf("invalid expression: %s, please check expression is valid!", s)
				return nil, errors.New(errMsg)
			}
		}

		if end > beg {
			if s[beg+1:end] == "#" {
				arr = append(arr, -1)
			} else {
				num, err := strconv.Atoi(s[beg+1 : end])
				if err != nil {
					return nil, err
				}
				arr = append(arr, num)
			}
			beg, end = -1, -1
		}
	}
	return arr, nil
}

func getOperands(s string) []string {
	r := strings.NewReplacer(
		"(", "",
		")", "",
		"+", " ",
		"-", " ",
		"*", " ",
		"/", " ")

	operands := strings.Split(r.Replace(s), " ")
	return operands
}
