package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func ParseColumnsSpec(spec string) ([][]string, error) {
	normalized := strings.ToUpper(strings.ReplaceAll(spec, " ", ""))
	if normalized == "" {
		return nil, fmt.Errorf("pole Stĺpce je povinné (napr. A-N,B-E,C-V)")
	}

	parts := strings.Split(normalized, ",")
	result := make([][]string, 0, len(parts))

	for _, part := range parts {
		split := strings.Split(part, "-")
		if len(split) != 2 || split[0] == "" || split[1] == "" {
			return nil, fmt.Errorf("neplatný formát stĺpcov: %q", part)
		}

		if split[1] != "N" && split[1] != "E" && split[1] != "V" {
			return nil, fmt.Errorf("neplatný typ %q v %q (povolené: N, E, V)", split[1], part)
		}

		result = append(result, split)
	}

	return result, nil
}

func ParsePositiveInt(value string, fieldName string) (int, error) {
	parsed, err := strconv.Atoi(strings.TrimSpace(value))
	if err != nil || parsed <= 0 {
		return 0, fmt.Errorf("%s musí byť kladné celé číslo", fieldName)
	}
	return parsed, nil
}

func ProcessWorkbook(inputPath, outputPath, sheet string, columnsSpec string, start, end int) error {
	if strings.TrimSpace(inputPath) == "" {
		return fmt.Errorf("vyber vstupný Excel súbor")
	}
	if strings.TrimSpace(outputPath) == "" {
		return fmt.Errorf("zvoľ cieľový súbor pre uloženie")
	}
	if strings.TrimSpace(sheet) == "" {
		return fmt.Errorf("vyber hárok")
	}
	if start > end {
		return fmt.Errorf("počiatočný riadok nesmie byť väčší ako koncový")
	}

	columns, err := ParseColumnsSpec(columnsSpec)
	if err != nil {
		return err
	}

	file, err := excelize.OpenFile(inputPath)
	if err != nil {
		return err
	}
	defer func() {
		_ = file.Close()
	}()

	exists := false
	for _, item := range file.GetSheetList() {
		if item == sheet {
			exists = true
			break
		}
	}
	if !exists {
		return fmt.Errorf("hárok %q sa v súbore nenašiel", sheet)
	}

	fm := NewFormatManager(file, sheet, start, end)
	if err := fm.FormatColumns(columns); err != nil {
		return err
	}

	return file.SaveAs(outputPath)
}
