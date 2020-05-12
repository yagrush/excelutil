package excelutil

import (
	"errors"
	"github.com/tealeg/xlsx"
	"math"
	"regexp"
	"strconv"
	"strings"
)

type ExcelFile struct {
	File     *xlsx.File
	FilePath string
}

// ex) excel := excelutil.ExcelFile{}.Init("./sample.xlsx")
func Init(excelFilePath string) (ret *ExcelFile, err error) {
	excelFileInstance := ExcelFile{}
	excelFileInstance.FilePath = excelFilePath
	excelFileInstance.File, err = xlsx.OpenFile(excelFileInstance.FilePath)
	return &excelFileInstance, err
}

func (excelFileInstance *ExcelFile) checkExcelOpenedAndExec(sheetName string, colNum, rowNum int,
	f func(sheetName string, colNum, rowNum int) (string, error)) (string, error) {
	if excelFileInstance == nil || excelFileInstance.File == nil {
		return "", errors.New("error: excelFile File is not loaded. call initial func")
	}
	return f(sheetName, colNum, rowNum)
}

// ex) excel.ReadCell("mysheet1", 3, 3) -> "hoge!"
func (excelFileInstance *ExcelFile) ReadCell(sheetName string, colNum, rowNum int) (string, error) {
	return excelFileInstance.checkExcelOpenedAndExec(sheetName, colNum, rowNum,
		func(sheetName string, colNum, rowNum int) (string, error) {
			if len(excelFileInstance.File.Sheet[sheetName].Rows) < rowNum {
				return "", errors.New("error: parameter rowNum is over")
			} else if len(excelFileInstance.File.Sheet[sheetName].Rows[rowNum-1].Cells) < colNum {
				return "", errors.New("error: parameter colNum is over")
			}
			return excelFileInstance.File.Sheet[sheetName].Rows[rowNum-1].Cells[colNum-1].String(), nil
		})
}

// ex) excel.ReadCellByCellAddress("mysheet1", "C3") -> "hoge!"
func (excelFileInstance *ExcelFile) ReadCellByCellAddress(sheetName string, cellAddress string) (string, error) {
	colNum, rowNum := ConvExcelCellAddressToColnumAndRownum(cellAddress)
	return excelFileInstance.ReadCell(sheetName, colNum, rowNum)
}

// ex) thisfunc("B:5") -> 2, 5
func ConvExcelCellAddressToColnumAndRownum(excelCellAddress string) (colnum, rownum int) {
	regex := `^([a-zA-Z]+)[^a-zA-Z0-9]*([\d]+)$`
	if !regexp.MustCompile(regex).Match([]byte(excelCellAddress)) {
		return colnum, rownum
	}
	sp := splitStringByRegexMatch(excelCellAddress, regex)

	rownum, err := strconv.Atoi(sp[1])
	if err != nil {
		return colnum, rownum
	}

	colnum = convExcelColAlphabetToNum(sp[0])
	if colnum == 0 {
		return colnum, rownum
	}
	return colnum, rownum
}

// ex) thisfunc(2) -> "B"
func ConvExcelColNumToAlphabet(cellColnum int) (excelColAlphabet string) {
	unitDigit := (cellColnum - 1) / 26
	unitsDigit := cellColnum - (unitDigit * 26)
	if unitDigit > 0 {
		excelColAlphabet = string(unitDigit + 64)
	}
	if unitsDigit > 0 {
		excelColAlphabet = excelColAlphabet + string(unitsDigit+64)
	}
	return excelColAlphabet
}

// ex) thisfunc("B") -> 2
func convExcelColAlphabetToNum(excelColAlphabet string) (ret int) {
	if !regexp.MustCompile(`^[a-zA-Z]+$`).Match([]byte(excelColAlphabet)) {
		return ret
	}
	for i, s := range strings.ToUpper(excelColAlphabet) {
		ret += int(math.Pow(26, float64(len(excelColAlphabet)-1-i))) * (int(s) - 64)
	}
	return ret
}

// ex) thisfunc("B:5", `^([a-zA-Z]+)[^a-zA-Z0-9]*([\d]+)$`) -> {"B", "5"}
func splitStringByRegexMatch(targetString, regex string) (ret []string) {
	bytesOfTargetString := []byte(targetString)
	regexMatchedBytesGroups := regexp.MustCompile(regex).FindSubmatch(bytesOfTargetString)
	ret = make([]string, len(regexMatchedBytesGroups)-1, len(regexMatchedBytesGroups)-1)
	for i, aRegexMatchedBytesGroup := range regexMatchedBytesGroups {
		if i == 0 {
			continue
		}
		ret[i-1] = string(aRegexMatchedBytesGroup)
	}
	return ret
}