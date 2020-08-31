package main

import (
	"github.com/tealeg/xlsx"
	"strconv"
	"time"
)

type Stock struct {
	Name string
	Code string
	NgwId string
}

func main(){
	UserSubStocks()
}


/**
	用户自选股票处理
	股票excel文件   stocks.xls    股票名称、股票代码、ngw_id
	用户excel文件   users.xls	 用户id、ngw_id、关注时间
 */
func UserSubStocks(){
	newUsers := xlsx.NewFile()
	stocksFile := "d:\\stocks.xlsx"
	usersFile := "d:\\users.xlsx"

	stockMap := map[string]*Stock{}
	//初始化stockMap
	stockXlFile, err := xlsx.OpenFile(stocksFile)
	if err != nil {
		panic(err.Error())
	}
	for _, curSheet := range stockXlFile.Sheets {
		for _, row := range curSheet.Rows {
			newStock := Stock{}
			newStock.Name = row.Cells[0].Value
			newStock.Code = row.Cells[1].Value
			newStock.NgwId = row.Cells[2].Value
			stockMap[newStock.NgwId] = &newStock
		}
	}

	userXlFile,err := xlsx.OpenFile(usersFile)
	if err != nil {
		panic(err.Error())
	}

	for i, curSheet := range userXlFile.Sheets {
		newSheet, err := newUsers.AddSheet("Sheet"+strconv.Itoa(i))
		if err != nil {
			panic(err.Error())
		}
		for _, row := range curSheet.Rows {
			newRow := newSheet.AddRow()
			for i:=0;i<6;i++{
				newRow.AddCell()
			}
			newRow.Cells[0].Value = row.Cells[0].Value
			newRow.Cells[1].Value = row.Cells[1].Value
			seconds,_ := strconv.Atoi(row.Cells[2].Value)
			newRow.Cells[2].Value = time.Unix(int64(seconds),0).Format("2006-01-02 15:04:05")
			if stock,ok := stockMap[row.Cells[1].Value];ok{
				newRow.Cells[3].Value = stock.Name
				newRow.Cells[4].Value = stock.Code
			}
		}
	}
	err = newUsers.Save("d:\\newUsers.xlsx")
	if err != nil {
		panic(err.Error())
	}
}
