//go:build windows

package main

import "github.com/sqweek/dialog"

func LoadFile() (string, bool, error) {
	path, err := dialog.File().
		Title("Vyber Excel súbor").
		Filter("Excel Files", "xls", "xlsx", "xlsm").
		Load()
	if err != nil {
		if err == dialog.ErrCancelled {
			return "", false, nil
		}
		return "", false, err
	}

	return path, true, nil
}

func SaveFile(defaultName string) (string, bool, error) {
	dlg := dialog.File().
		Title("Uložiť ako").
		Filter("Excel Files", "xlsx")
	if defaultName != "" {
		dlg = dlg.SetStartFile(defaultName)
	}

	path, err := dlg.Save()
	if err != nil {
		if err == dialog.ErrCancelled {
			return "", false, nil
		}
		return "", false, err
	}

	return path, true, nil
}
