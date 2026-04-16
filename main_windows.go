//go:build windows

package main

import (
	"strings"

	"github.com/sqweek/dialog"
)

func main() {
	config, cancelled, err := ShowMainWindow(version)
	if err != nil {
		showError(err)
		return
	}
	if cancelled {
		return
	}

	if err := ProcessWorkbook(config.InputPath, config.OutputPath, config.Sheet, config.Columns, config.StartRow, config.EndRow); err != nil {
		showError(err)
		return
	}

	dialog.Message("Hotovo. Súbor bol úspešne uložený:\n%s", config.OutputPath).Title("Úspech").Info()
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
