//go:build windows

package main

import (
	"fmt"
	"strings"

	"github.com/sqweek/dialog"
)

func main() {
	dialog.Message("Excel Súradnice Script %s\n\nTento sprievodca ťa prevedie celým formátovaním.", version).
		Title("Excel Súradnice Script").
		Info()

	inputPath, ok, err := LoadFile()
	if err != nil {
		showError(err)
		return
	}
	if !ok {
		return
	}

	file, err := openWorkbook(inputPath)
	if err != nil {
		showError(err)
		return
	}
	sheets := file.GetSheetList()
	_ = file.Close()
	if len(sheets) == 0 {
		showError(fmt.Errorf("zošit neobsahuje žiadne hárky"))
		return
	}

	sheet, err := PromptChoice("Vyber hárok", "Hárok", sheets)
	if err != nil {
		showError(err)
		return
	}

	columns, err := PromptText("Stĺpce", "Zadaj stĺpce (napr. A-N,B-E,C-V):", "A-N")
	if err != nil {
		showError(err)
		return
	}

	startStr, err := PromptText("Od riadku", "Zadaj počiatočný riadok:", "1")
	if err != nil {
		showError(err)
		return
	}
	start, err := ParsePositiveInt(startStr, "Od riadku")
	if err != nil {
		showError(err)
		return
	}

	endStr, err := PromptText("Do riadku", "Zadaj koncový riadok:", startStr)
	if err != nil {
		showError(err)
		return
	}
	end, err := ParsePositiveInt(endStr, "Do riadku")
	if err != nil {
		showError(err)
		return
	}

	outputPath, ok, err := SaveFile(defaultOutputPath(inputPath))
	if err != nil {
		showError(err)
		return
	}
	if !ok {
		return
	}

	summary := fmt.Sprintf("Vstup: %s\nHárok: %s\nStĺpce: %s\nRiadky: %d-%d\nVýstup: %s", inputPath, sheet, columns, start, end, outputPath)
	if !dialog.Message("Skontroluj nastavenia:\n\n%s\n\nPokračovať?", summary).Title("Potvrdenie").YesNo() {
		return
	}

	if err := ProcessWorkbook(inputPath, outputPath, sheet, columns, start, end); err != nil {
		showError(err)
		return
	}

	dialog.Message("Hotovo. Súbor bol úspešne uložený:\n%s", outputPath).Title("Úspech").Info()
}

func defaultOutputPath(input string) string {
	if input == "" {
		return ""
	}
	lower := strings.ToLower(input)
	if strings.HasSuffix(lower, ".xlsx") {
		return input[:len(input)-5] + "_formatted.xlsx"
	}
	if strings.HasSuffix(lower, ".xlsm") || strings.HasSuffix(lower, ".xls") {
		return input + "_formatted.xlsx"
	}
	return input + "_formatted.xlsx"
}

func showError(err error) {
	if err == nil {
		return
	}
	dialog.Message("%v", err).Title("Chyba").Error()
}
