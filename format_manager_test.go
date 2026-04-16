package main

import (
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestFormatColumnsDeletesInvalidCoordinateRows(t *testing.T) {
	file := excelize.NewFile()
	defer func() {
		_ = file.Close()
	}()

	sheet := file.GetSheetName(file.GetActiveSheetIndex())
	_ = file.SetCellStr(sheet, "A1", "48123")
	_ = file.SetCellStr(sheet, "A2", "x")
	_ = file.SetCellStr(sheet, "A3", "49123")

	fm := NewFormatManager(file, sheet, 1, 3)
	if err := fm.FormatColumns([][]string{{"A", "N"}}); err != nil {
		t.Fatalf("FormatColumns returned error: %v", err)
	}

	got1, _ := file.GetCellValue(sheet, "A1")
	got2, _ := file.GetCellValue(sheet, "A2")

	if got1 != "48N123" {
		t.Fatalf("expected first value to be formatted, got %q", got1)
	}
	if got2 != "49N123" {
		t.Fatalf("expected invalid row to be deleted, got %q", got2)
	}
}

func TestFormatColumnsFormatsEastingAndHeight(t *testing.T) {
	file := excelize.NewFile()
	defer func() {
		_ = file.Close()
	}()

	sheet := file.GetSheetName(file.GetActiveSheetIndex())
	_ = file.SetCellStr(sheet, "A1", "48N123")
	_ = file.SetCellStr(sheet, "B1", "123.6m")

	fm := NewFormatManager(file, sheet, 1, 1)
	if err := fm.FormatColumns([][]string{{"A", "E"}, {"B", "V"}}); err != nil {
		t.Fatalf("FormatColumns returned error: %v", err)
	}

	gotCoordinate, _ := file.GetCellValue(sheet, "A1")
	gotHeight, _ := file.GetCellValue(sheet, "B1")

	if gotCoordinate != "48E123" {
		t.Fatalf("expected coordinate to be converted to easting, got %q", gotCoordinate)
	}
	if gotHeight != "123" {
		t.Fatalf("expected height to be normalized, got %q", gotHeight)
	}
}
