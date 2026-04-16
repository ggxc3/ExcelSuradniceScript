//go:build windows

package main

import "testing"

func TestDefaultOutputPath(t *testing.T) {
	tests := []struct {
		name  string
		input string
		want  string
	}{
		{name: "empty", input: "", want: ""},
		{name: "xlsx", input: `C:\\tmp\\input.xlsx`, want: `C:\\tmp\\input_formatted.xlsx`},
		{name: "xlsx uppercase", input: `C:\\tmp\\INPUT.XLSX`, want: `C:\\tmp\\INPUT_formatted.xlsx`},
		{name: "xlsm", input: `C:\\tmp\\input.xlsm`, want: `C:\\tmp\\input.xlsm_formatted.xlsx`},
		{name: "xls", input: `C:\\tmp\\input.xls`, want: `C:\\tmp\\input.xls_formatted.xlsx`},
		{name: "other extension", input: `C:\\tmp\\input.csv`, want: `C:\\tmp\\input.csv_formatted.xlsx`},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := defaultOutputPath(tt.input)
			if got != tt.want {
				t.Fatalf("defaultOutputPath(%q) = %q, want %q", tt.input, got, tt.want)
			}
		})
	}
}
