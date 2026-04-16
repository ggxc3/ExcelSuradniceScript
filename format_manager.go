package main

import (
	"fmt"
	"math"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type FormatManager struct {
	file  *excelize.File
	sheet string
	start int
	end   int
}

func NewFormatManager(file *excelize.File, sheet string, start, end int) *FormatManager {
	return &FormatManager{
		file:  file,
		sheet: sheet,
		start: start,
		end:   end,
	}
}

func (fm *FormatManager) FormatColumns(columns [][]string) error {
	for _, col := range columns {
		if err := fm.formatColumn(col[0], col[1]); err != nil {
			return err
		}
	}

	return nil
}

func (fm *FormatManager) formatColumn(column, coorType string) error {
	deleteRows := make([][2]int, 0)
	falsing := false

	for i := fm.start; i <= fm.end; i++ {
		fmt.Printf("%s-%s: Riadok: %d\n", column, coorType, i)
		if coorType == "V" {
			if err := fm.formatHeightCell(column, i); err != nil {
				return err
			}
			continue
		}

		ok, err := fm.formatCell(column, i, coorType)
		if err != nil {
			return err
		}

		if !ok {
			if !falsing {
				deleteRows = append(deleteRows, [2]int{i, 1})
				falsing = true
			} else {
				deleteRows[len(deleteRows)-1][1]++
			}
		} else if falsing {
			falsing = false
		}
	}

	index := 1
	for i := len(deleteRows) - 1; i >= 0; i-- {
		fmt.Printf("Mažem nespravne riadky %d/%d\n", index, len(deleteRows))
		for deleted := 0; deleted < deleteRows[i][1]; deleted++ {
			if err := fm.file.RemoveRow(fm.sheet, deleteRows[i][0]); err != nil {
				return err
			}
		}
		fm.end -= deleteRows[i][1]
		index++
	}

	return nil
}

func (fm *FormatManager) formatCell(col string, row int, coorType string) (bool, error) {
	cell := fmt.Sprintf("%s%d", col, row)
	text, err := fm.file.GetCellValue(fm.sheet, cell)
	if err != nil {
		return false, err
	}

	if text == "" || len(text) < 5 || text[1] == '.' || text[1] == ',' {
		return false, nil
	}

	text = strings.ReplaceAll(text, " ", "")
	text = strings.ReplaceAll(text, ",", ".")

	if text[0] == '0' {
		text = text[1:]
	}

	if startsWithTwoDigits(text) || ((text[0] == 'E' || text[0] == 'N') && isDigit(text[1])) {
		if coorType == "E" {
			text = strings.ReplaceAll(text, "N", "")
			if !strings.Contains(text, "E") {
				text = text[:2] + "E" + text[2:]
			}
		} else if coorType == "N" {
			text = strings.ReplaceAll(text, "E", "")
			if !strings.Contains(text, "N") {
				text = text[:2] + "N" + text[2:]
			}
		}
	} else {
		return false, nil
	}

	return true, fm.file.SetCellStr(fm.sheet, cell, text)
}

func (fm *FormatManager) formatHeightCell(col string, row int) error {
	cell := fmt.Sprintf("%s%d", col, row)
	text, err := fm.file.GetCellValue(fm.sheet, cell)
	if err != nil {
		return err
	}

	if text == "" {
		return nil
	}

	text = strings.ReplaceAll(text, " ", "")
	text = strings.ReplaceAll(text, ".", ",")

	normalizedForParse := strings.ReplaceAll(text, ",", ".")
	if _, err := strconv.ParseFloat(normalizedForParse, 64); err != nil {
		var result strings.Builder
		for _, item := range text {
			if item >= '0' && item <= '9' {
				result.WriteRune(item)
				continue
			}

			return fm.file.SetCellStr(fm.sheet, cell, result.String())
		}
		text = result.String()
	} else {
		value, _ := strconv.ParseFloat(normalizedForParse, 64)
		text = strconv.FormatInt(int64(math.Round(value)), 10)
	}

	text = strings.ReplaceAll(text, ",", ".")
	return fm.file.SetCellStr(fm.sheet, cell, text)
}

func startsWithTwoDigits(text string) bool {
	return len(text) >= 2 && isDigit(text[0]) && isDigit(text[1])
}

func isDigit(char byte) bool {
	return char >= '0' && char <= '9'
}
