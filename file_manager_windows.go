//go:build windows

package main

import "github.com/sqweek/dialog"

func LoadFile() (string, bool, error) {
	path, err := dialog.File().
		Title("Vyber Excel súbor").
		Filter("Excel Files", "xls", "xlsx", "xlsm").
		Load()
	if err != nil {
		if err == dialog.Cancelled {
			return "", false, nil
		}
		return "", false, err
	}

	return path, true, nil
}

func SaveFile() (string, bool, error) {
	path, err := dialog.File().
		Title("Uložiť ako").
		Filter("Excel Files", "xls", "xlsx", "xlsm").
		SetStartFile("output.xlsx").
		Save()
	if err != nil {
		if err == dialog.Cancelled {
			return "", false, nil
		}
		return "", false, err
	}

	return path, true, nil
}
