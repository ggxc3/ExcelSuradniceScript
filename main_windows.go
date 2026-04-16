//go:build windows

package main

import (
	"fmt"
	"strings"
	"syscall"
	"unsafe"

	"github.com/TheTitanrain/w32"
)

const (
	idInputEdit  = 101
	idInputBtn   = 102
	idSheetEdit  = 103
	idSheetHint  = 104
	idColsEdit   = 105
	idStartEdit  = 106
	idEndEdit    = 107
	idOutputEdit = 108
	idOutputBtn  = 109
	idRunBtn     = 110
	idStatus     = 111
)

type uiState struct {
	hwndMain    w32.HWND
	hInputEdit  w32.HWND
	hSheetEdit  w32.HWND
	hSheetHint  w32.HWND
	hColsEdit   w32.HWND
	hStartEdit  w32.HWND
	hEndEdit    w32.HWND
	hOutputEdit w32.HWND
	hRunBtn     w32.HWND
	hStatus     w32.HWND
}

var ui uiState

func main() {
	if err := runDesktopWindow(); err != nil {
		w32.MessageBox(0, err.Error(), "Chyba", w32.MB_OK|w32.MB_ICONERROR)
	}
}

func runDesktopWindow() error {
	hInstance := w32.GetModuleHandle("")
	className := syscall.StringToUTF16Ptr("ExcelSuradniceDesktopWindow")

	wc := w32.WNDCLASSEX{
		Size:       uint32(unsafe.Sizeof(w32.WNDCLASSEX{})),
		WndProc:    syscall.NewCallback(windowProc),
		Instance:   hInstance,
		Cursor:     w32.LoadCursor(0, (*uint16)(unsafe.Pointer(uintptr(w32.IDC_ARROW)))),
		Background: w32.HBRUSH(w32.GetStockObject(w32.WHITE_BRUSH)),
		ClassName:  className,
	}
	if w32.RegisterClassEx(&wc) == 0 {
		return fmt.Errorf("nepodarilo sa zaregistrovať okno")
	}

	title := fmt.Sprintf("Excel Súradnice Script %s", version)
	hwnd := w32.CreateWindowEx(
		0,
		className,
		syscall.StringToUTF16Ptr(title),
		w32.WS_OVERLAPPEDWINDOW|w32.WS_VISIBLE,
		w32.CW_USEDEFAULT,
		w32.CW_USEDEFAULT,
		860,
		560,
		0,
		0,
		hInstance,
		nil,
	)
	if hwnd == 0 {
		return fmt.Errorf("nepodarilo sa vytvoriť hlavné okno")
	}

	ui.hwndMain = hwnd
	w32.ShowWindow(hwnd, w32.SW_SHOWDEFAULT)
	w32.UpdateWindow(hwnd)

	var msg w32.MSG
	for w32.GetMessage(&msg, 0, 0, 0) > 0 {
		w32.TranslateMessage(&msg)
		w32.DispatchMessage(&msg)
	}

	return nil
}

func windowProc(hwnd w32.HWND, msg uint32, wParam, lParam uintptr) uintptr {
	switch msg {
	case w32.WM_CREATE:
		buildLayout(hwnd)
		return 0
	case w32.WM_COMMAND:
		handleCommand(hwnd, loword(wParam))
		return 0
	case w32.WM_DESTROY:
		w32.PostQuitMessage(0)
		return 0
	default:
		return w32.DefWindowProc(hwnd, msg, wParam, lParam)
	}
}

func buildLayout(hwnd w32.HWND) {
	font := w32.GetStockObject(w32.DEFAULT_GUI_FONT)

	createStatic(hwnd, "Desktop formátovanie Excel súradníc a výšky", 20, 20, 500, 24, font)
	createStatic(hwnd, "Stĺpce zapisuj ako: A-N,B-E,C-V", 20, 45, 360, 20, font)

	createStatic(hwnd, "Vstupný súbor:", 20, 80, 140, 20, font)
	ui.hInputEdit = createEdit(hwnd, "", 20, 102, 620, 26, true, idInputEdit, font)
	createButton(hwnd, "Vybrať...", 660, 102, 160, 26, idInputBtn, font)

	createStatic(hwnd, "Hárok:", 20, 140, 80, 20, font)
	ui.hSheetEdit = createEdit(hwnd, "", 20, 162, 220, 26, false, idSheetEdit, font)
	ui.hSheetHint = createStatic(hwnd, "Dostupné hárky: -", 260, 166, 560, 20, font)

	createStatic(hwnd, "Stĺpce:", 20, 200, 80, 20, font)
	ui.hColsEdit = createEdit(hwnd, "A-N", 20, 222, 260, 26, false, idColsEdit, font)

	createStatic(hwnd, "Od riadku:", 300, 200, 80, 20, font)
	ui.hStartEdit = createEdit(hwnd, "1", 300, 222, 100, 26, false, idStartEdit, font)

	createStatic(hwnd, "Do riadku:", 420, 200, 80, 20, font)
	ui.hEndEdit = createEdit(hwnd, "1", 420, 222, 100, 26, false, idEndEdit, font)

	createStatic(hwnd, "Výstupný súbor:", 20, 265, 140, 20, font)
	ui.hOutputEdit = createEdit(hwnd, "", 20, 287, 620, 26, true, idOutputEdit, font)
	createButton(hwnd, "Uložiť ako...", 660, 287, 160, 26, idOutputBtn, font)

	ui.hRunBtn = createButton(hwnd, "Spracovať", 20, 340, 180, 36, idRunBtn, font)
	ui.hStatus = createStatic(hwnd, "Pripravené. Vyber vstupný súbor.", 220, 349, 600, 20, font)
}

func handleCommand(hwnd w32.HWND, controlID uint16) {
	switch controlID {
	case idInputBtn:
		onChooseInput(hwnd)
	case idOutputBtn:
		onChooseOutput(hwnd)
	case idRunBtn:
		onProcess(hwnd)
	}
}

func onChooseInput(hwnd w32.HWND) {
	path, ok, err := LoadFile()
	if err != nil {
		showError(err)
		return
	}
	if !ok {
		return
	}

	w32.SetWindowText(ui.hInputEdit, path)
	w32.SetWindowText(ui.hOutputEdit, defaultOutputPath(path))

	file, err := openWorkbook(path)
	if err != nil {
		showError(err)
		return
	}
	defer func() { _ = file.Close() }()

	sheets := file.GetSheetList()
	if len(sheets) == 0 {
		showError(fmt.Errorf("zošit neobsahuje žiadne hárky"))
		return
	}

	w32.SetWindowText(ui.hSheetEdit, sheets[0])
	hint := "Dostupné hárky: " + strings.Join(sheets, ", ")
	w32.SetWindowText(ui.hSheetHint, hint)
	w32.SetWindowText(ui.hStatus, "Súbor načítaný. Skontroluj nastavenia a klikni Spracovať.")

	_ = hwnd
}

func onChooseOutput(hwnd w32.HWND) {
	defaultName := w32.GetWindowText(ui.hOutputEdit)
	if defaultName == "" {
		defaultName = defaultOutputPath(w32.GetWindowText(ui.hInputEdit))
	}
	path, ok, err := SaveFile(defaultName)
	if err != nil {
		showError(err)
		return
	}
	if !ok {
		return
	}

	w32.SetWindowText(ui.hOutputEdit, path)
	_ = hwnd
}

func onProcess(hwnd w32.HWND) {
	inputPath := strings.TrimSpace(w32.GetWindowText(ui.hInputEdit))
	outputPath := strings.TrimSpace(w32.GetWindowText(ui.hOutputEdit))
	sheet := strings.TrimSpace(w32.GetWindowText(ui.hSheetEdit))
	columns := strings.TrimSpace(w32.GetWindowText(ui.hColsEdit))
	startStr := strings.TrimSpace(w32.GetWindowText(ui.hStartEdit))
	endStr := strings.TrimSpace(w32.GetWindowText(ui.hEndEdit))

	start, err := ParsePositiveInt(startStr, "Od riadku")
	if err != nil {
		showError(err)
		return
	}
	end, err := ParsePositiveInt(endStr, "Do riadku")
	if err != nil {
		showError(err)
		return
	}

	w32.EnableWindow(ui.hRunBtn, false)
	w32.SetWindowText(ui.hStatus, "Spracovávam súbor, čakaj prosím...")

	err = ProcessWorkbook(inputPath, outputPath, sheet, columns, start, end)
	if err != nil {
		w32.EnableWindow(ui.hRunBtn, true)
		showError(err)
		w32.SetWindowText(ui.hStatus, "Chyba spracovania. Skontroluj vstupy.")
		return
	}

	w32.EnableWindow(ui.hRunBtn, true)
	w32.SetWindowText(ui.hStatus, "Hotovo. Súbor bol úspešne uložený.")
	w32.MessageBox(hwnd, "Formátovanie bolo dokončené.", "Úspech", w32.MB_OK|w32.MB_ICONINFORMATION)
}

func createStatic(parent w32.HWND, text string, x, y, width, height int, font w32.HGDIOBJ) w32.HWND {
	h := w32.CreateWindowEx(
		0,
		syscall.StringToUTF16Ptr("STATIC"),
		syscall.StringToUTF16Ptr(text),
		w32.WS_CHILD|w32.WS_VISIBLE|w32.SS_LEFT,
		x, y, width, height,
		parent,
		0,
		0,
		nil,
	)
	w32.SendMessage(h, w32.WM_SETFONT, uintptr(font), 1)
	return h
}

func createEdit(parent w32.HWND, text string, x, y, width, height int, readOnly bool, id uintptr, font w32.HGDIOBJ) w32.HWND {
	style := uint(w32.WS_CHILD | w32.WS_VISIBLE | w32.WS_BORDER | w32.ES_AUTOHSCROLL | w32.WS_TABSTOP)
	if readOnly {
		style |= w32.ES_READONLY
	}

	h := w32.CreateWindowEx(
		w32.WS_EX_CLIENTEDGE,
		syscall.StringToUTF16Ptr("EDIT"),
		syscall.StringToUTF16Ptr(text),
		style,
		x, y, width, height,
		parent,
		w32.HMENU(id),
		0,
		nil,
	)
	w32.SendMessage(h, w32.WM_SETFONT, uintptr(font), 1)
	return h
}

func createButton(parent w32.HWND, text string, x, y, width, height int, id uintptr, font w32.HGDIOBJ) w32.HWND {
	h := w32.CreateWindowEx(
		0,
		syscall.StringToUTF16Ptr("BUTTON"),
		syscall.StringToUTF16Ptr(text),
		w32.WS_CHILD|w32.WS_VISIBLE|w32.WS_TABSTOP,
		x, y, width, height,
		parent,
		w32.HMENU(id),
		0,
		nil,
	)
	w32.SendMessage(h, w32.WM_SETFONT, uintptr(font), 1)
	return h
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
	w32.MessageBox(0, err.Error(), "Chyba", w32.MB_OK|w32.MB_ICONERROR)
}

func loword(value uintptr) uint16 {
	return uint16(value & 0xFFFF)
}
