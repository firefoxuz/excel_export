package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"math/rand"
	"strconv"
)

const letterBytes = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

func RandStringBytes(n int) string {
	b := make([]byte, n)
	for i := range b {
		b[i] = letterBytes[rand.Intn(len(letterBytes))]
	}
	return string(b)
}

func main() {
	fmt.Print("Starting ...")

	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// Create a new sheet.
	index, err := f.NewSheet("Sheet1")
	if err != nil {
		fmt.Println(err)
		return
	}

	for i := 0; i < 1000000; i++ {
		rowNumber := strconv.Itoa(i)
		f.SetCellValue("Sheet1", "A"+rowNumber, RandStringBytes(10))
		f.SetCellValue("Sheet1", "B"+rowNumber, RandStringBytes(10))
		f.SetCellValue("Sheet1", "C"+rowNumber, RandStringBytes(10))
		f.SetCellValue("Sheet1", "D"+rowNumber, RandStringBytes(10))
		f.SetCellValue("Sheet1", "E"+rowNumber, RandStringBytes(10))
		f.SetCellValue("Sheet1", "F"+rowNumber, RandStringBytes(10))
		f.SetCellValue("Sheet1", "G"+rowNumber, RandStringBytes(10))
	}

	f.SetActiveSheet(index)
	if err := f.SaveAs("test.xlsx"); err != nil {
		fmt.Println(err)
	}
}
