package main

import (
	"errors"
	"flag"
	"fmt"
	"os"
    "log"
    "encoding/csv"

	"github.com/tealeg/xlsx"
)

var xlsxPath = flag.String("f", "", "Path to an XLSX file")
var sheetIndex = flag.Int("i", 0, "Index of sheet to convert, zero based")
var delimiter = flag.String("d", ";", "Delimiter to use between fields")
var maxRows   = flag.Int("r", 1000, "max number of rows to process")

func generateCSVFromXLSXFile(excelFileName string, sheetIndex int) error {
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		return err
	}
	sheetLen := len(xlFile.Sheets)
	switch {
	case sheetLen == 0:
		return errors.New("This XLSX file contains no sheets")
	case sheetIndex >= sheetLen:
		return fmt.Errorf("No sheet %d available, please select a sheet between 0 and %d", sheetIndex, sheetLen-1)
	}
	sheet := xlFile.Sheets[sheetIndex]

    w := csv.NewWriter(os.Stdout)
	for i, row := range sheet.Rows {
        if i >= *maxRows {
            break
        }

		var vals []string
		if row != nil {
			for _, cell := range row.Cells {
				str, err := cell.FormattedValue()
				if err != nil {
					vals = append(vals, err.Error())
				} else {
                    vals = append(vals, str)
                }
			}
            w.Write(vals)
		}
	}
	return nil
}

func main() {
	flag.Parse()
	if len(os.Args) < 3 {
		flag.PrintDefaults()
		return
	}
	flag.Parse()
	if err := generateCSVFromXLSXFile(*xlsxPath, *sheetIndex); err != nil {
        log.Println(err)
	}
}
