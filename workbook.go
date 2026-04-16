package main

import "github.com/xuri/excelize/v2"

func openWorkbook(path string) (*excelize.File, error) {
	return excelize.OpenFile(path)
}
