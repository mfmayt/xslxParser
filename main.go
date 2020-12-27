package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"io"
	"os"

	"github.com/tealeg/xlsx/v3"
)

// Add column values here
const enColumn = "Value_ENG"
const trColumn = "Value_TR"
const keyColumn = "Key"

func check(e error) {
	if e != nil {
		panic(e)
	}
}

func xlsxToJSON() {
	// Change name of your file
	filename := "sample.xlsx"
	wb, err := xlsx.OpenFile(filename)
	check(err)

	rows := make([][]map[string]string, len(wb.Sheets))
	// add other languages to here
	languageColumns := []string{enColumn, trColumn}

	for _, foldername := range languageColumns {
		if _, err := os.Stat(foldername); os.IsNotExist(err) {
			os.Mkdir(foldername, 0777)
		}
		check(err)
	}

	for sheetIteration, sh := range wb.Sheets {
		columns := make(map[string]int)

		for col := 0; col < sh.MaxCol; col++ {
			cell, err := sh.Cell(0, col)
			check(err)
			value, err := cell.FormattedValue()
			check(err)
			if value != "" {
				columns[value] = col
			}
		}
		rows[sheetIteration] = make([]map[string]string, len(languageColumns))
		for langIteration, lang := range languageColumns {
			rows[sheetIteration][langIteration] = make(map[string]string)
			for row := 1; row < sh.MaxRow; row++ {
				valueCell, err := sh.Cell(row, columns[lang])
				keyCell, err := sh.Cell(row, columns[keyColumn])
				val, err := valueCell.FormattedValue()
				key, err := keyCell.FormattedValue()
				if err != nil {
					panic(err)
				}
				if val != "" && key != "" {
					rows[sheetIteration][langIteration][key] = val
				}
			}

			var toJSONValues = rows[sheetIteration][langIteration]
			f, err := os.Create(lang + "/" + sh.Name)
			w := bufio.NewWriter(f)
			check(err)
			ToJSON(toJSONValues, w)
			w.Flush()
			check(err)
		}
	}
}

func main() {
	fmt.Println("== xlsx package tutorial ==")
	xlsxToJSON()
}

// ToJSON serializes the given interface into a string based JSON format
func ToJSON(i interface{}, w io.Writer) error {
	e := json.NewEncoder(w)
	e.SetIndent("", "\t")

	return e.Encode(i)
}
