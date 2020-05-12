package main

import (
	"github.com/yagrush/excelutil"
	"log"
)

func main() {
	excel, err := excelutil.Init("./sample.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	log.Println(excelutil.ConvExcelCellAddressToColnumAndRownum("C2"))
	log.Println(excelutil.ConvExcelCellAddressToColnumAndRownum("AK43"))
	log.Println(excel.ReadCell("mysheet1", 3, 2))
	log.Println(excel.ReadCellByCellAddress("mysheet1", "C2"))
}
