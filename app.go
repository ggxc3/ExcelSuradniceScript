package main

import (
	"bufio"
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type App struct {
	file       *excelize.File
	mainSheet  string
	fileLoaded bool
	reader     *bufio.Reader
}

func NewApp(reader *bufio.Reader) (*App, error) {
	app := &App{reader: reader}

	pathToFile, ok, err := LoadFile()
	if err != nil {
		return app, err
	}
	if !ok {
		fmt.Println("Chyba pri načítavaní súboru.")
		return app, nil
	}

	fmt.Println("Načítavám súbor.")
	file, err := excelize.OpenFile(pathToFile)
	if err != nil {
		return app, err
	}

	app.file = file
	app.fileLoaded = true

	fmt.Println("Súbor úspešne načítaný.")
	fmt.Println("Počkaj na ďalší pokyn.")

	sheets := app.file.GetSheetList()
	if len(sheets) == 0 {
		return app, fmt.Errorf("zošit neobsahuje žiadne hárky")
	}

	app.mainSheet = sheets[app.selectSheet(sheets)-1]
	return app, nil
}

func (a *App) selectSheet(sheets []string) int {
	number := 1
	for _, sheet := range sheets {
		fmt.Printf("%d. %s\n", number, sheet)
		number++
	}

	for {
		input, _ := a.readLine()
		selected, err := strconv.Atoi(input)
		if err != nil {
			fmt.Println("Zadaj cislo: ")
			continue
		}

		if selected >= 1 && selected <= len(sheets) {
			return selected
		}

		fmt.Println("Zadaj spravne cislo z rozsahu: ")
	}
}

func (a *App) selectCols() [][]string {
	for {
		fmt.Println("Zadaj stlpce na fomatovanie (oddelene ciarkou): ")
		stringCols, _ := a.readLine()

		normalized := strings.ToUpper(strings.ReplaceAll(stringCols, " ", ""))
		parts := strings.Split(normalized, ",")

		invalid := normalized == ""
		if !invalid {
			for _, part := range parts {
				split := strings.Split(part, "-")
				if len(split) == 1 {
					invalid = true
					break
				}

				if split[1] != "N" && split[1] != "E" && split[1] != "V" {
					invalid = true
					break
				}
			}
		}

		if invalid {
			continue
		}

		result := make([][]string, len(parts))
		for i, part := range parts {
			result[i] = strings.Split(part, "-")
		}
		return result
	}
}

func (a *App) selectStartNumber() int {
	for {
		fmt.Println("Zadaj startovacie cislo riadka: ")
		input, _ := a.readLine()
		number, err := strconv.Atoi(input)
		if err == nil {
			return number
		}
	}
}

func (a *App) selectEndNumber() int {
	for {
		fmt.Println("Zadaj konciace cislo riadka: ")
		input, _ := a.readLine()
		number, err := strconv.Atoi(input)
		if err == nil {
			return number
		}
	}
}

func (a *App) saveAs() error {
	pathToFile, ok, err := SaveFile()
	if err != nil {
		return err
	}
	if !ok {
		fmt.Println("Chyba pri ulkadaní súboru.")
		return nil
	}

	return a.file.SaveAs(pathToFile)
}

func (a *App) Start() error {
	defer func() {
		if a.file != nil {
			_ = a.file.Close()
		}
	}()

	if !a.fileLoaded {
		return nil
	}

	cols := a.selectCols()
	start := a.selectStartNumber()
	end := a.selectEndNumber()

	fm := NewFormatManager(a.file, a.mainSheet, start, end)
	if err := fm.FormatColumns(cols); err != nil {
		return err
	}

	fmt.Println("Ukladám súbor.")
	return a.saveAs()
}

func (a *App) readLine() (string, error) {
	line, err := a.reader.ReadString('\n')
	if err != nil && len(line) == 0 {
		return "", err
	}

	return strings.TrimSpace(line), nil
}
