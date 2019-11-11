package main

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strconv"
	"time"
)

func main(){
	var newFile *xlsx.File
	var newSheet *xlsx.Sheet
	var newRow *xlsx.Row
	var newCell *xlsx.Cell

	newFile = xlsx.NewFile()

	oldXlsFile := "d:\\act.xlsx"
	xlFile, err := xlsx.OpenFile(oldXlsFile)
	if err != nil {
		fmt.Println(err.Error())
	}
	for i, curSheet := range xlFile.Sheets {
		newSheet, err = newFile.AddSheet("Sheet"+strconv.Itoa(i))
		if err != nil {
			fmt.Printf(err.Error())
		}
		for _, row := range curSheet.Rows {
			newRow = newSheet.AddRow()
			for i, cell := range row.Cells {
				newCell = newRow.AddCell()
				if i == 7{
					seconds,_ := strconv.Atoi(cell.Value)
					newCell.Value = time.Unix(int64(seconds),0).Format("2006-01-02 15:04:05")
				}else{
					newCell.Value = cell.Value
				}
			}
		}
	}
	err = newFile.Save("d:\\双十一.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}
}
