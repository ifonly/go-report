package xlsxExport

import (
	"fmt"
	"github.com/ifonly/go-report/model"
	"github.com/tealeg/xlsx"
)

type Export struct {
	SheetName         string
	TitleName         string
	FieldKeys         []*model.FieldKey
	FieldValues       []*model.FieldValue
	File              string
	TitleStyle        *xlsx.Style
	KeyFirstRowStyle  *xlsx.Style
	KeySecondRowStyle *xlsx.Style
	ValueStyle        *xlsx.Style
}

func (export Export) Save() {
	file := xlsx.NewFile()
	sheet, err := file.AddSheet(export.SheetName)
	if err != nil {
		fmt.Printf(err.Error())
	}
	titleRow := sheet.AddRow()
	titleCell := titleRow.AddCell()
	titleCell.Value = export.TitleName

	createKeyRow(sheet, export)
	titleCell.HMerge = sheet.MaxCol - 1

	titleCell.SetStyle(export.TitleStyle)

	createValueRow(sheet, export)

	err = file.Save(export.File)
	if err != nil {
		fmt.Printf(err.Error())
	}

	fmt.Println("\n\nexport success")
}

func createKeyRow(sheet *xlsx.Sheet, export Export) {
	fieldKeys := export.FieldKeys
	firstKeyRow := sheet.AddRow()
	currCell := 0
	serialCell := firstKeyRow.AddCell()
	serialCell.Value = "序号"
	serialCell.SetStyle(export.KeyFirstRowStyle)
	for _, fieldKey := range fieldKeys {
		firstKeyRowCell := firstKeyRow.AddCell()
		firstKeyRowCell.Value = fieldKey.Name
		firstKeyRowCell.SetStyle(export.KeyFirstRowStyle)
		// autoSizeColumn
		childrenKeys := fieldKey.ChildrenList
		childrenSize := len(childrenKeys)
		for i := 1; i < childrenSize; i++ {
			firstKeyRow.AddCell()
		}

		if childrenSize > 0 {
			firstKeyRowCell.HMerge = (childrenSize - 1)
			rowSize := len(sheet.Rows)
			secondRowKeyRow := sheet.Row(2)

			// 第二列增加空序号列
			if rowSize == 2 {
				secondRowKeyRow.AddCell()
			}

			for j := 0; j < currCell; j++ {
				secondRowKeyRow.AddCell()
			}
			for _, childKey := range childrenKeys {
				secondCell := secondRowKeyRow.AddCell()
				secondCell.Value = childKey.Name
				secondCell.SetStyle(export.KeySecondRowStyle)
			}
			currCell = 0
		} else {
			currCell++
			serialCell.VMerge = 1
			firstKeyRowCell.VMerge = 1
		}
	}
}

func createValueRow(sheet *xlsx.Sheet, export Export) {
	fieldValues := export.FieldValues
	for currValueRow, fieldValue := range fieldValues {
		fieldValueRow := sheet.AddRow()
		serialCell := fieldValueRow.AddCell()
		serialCell.SetInt(currValueRow + 1)

		rowValues := fieldValue.ChildrenList
		maxRow := sheet.MaxRow - 1
		mergeMaxRowSize := 0
		for _, cellValue := range rowValues {
			childrenFieldValues := cellValue.ChildrenList
			rowSize := len(childrenFieldValues)
			if mergeMaxRowSize < rowSize {
				mergeMaxRowSize = rowSize
			}
		}
		// 减去第一行
		mergeMaxRowSize = mergeMaxRowSize - 1
		serialCell.VMerge = mergeMaxRowSize

		// 当前列
		currCell := 0
		for _, cellValue := range rowValues {
			childrenFieldValues := cellValue.ChildrenList

			if len(childrenFieldValues) > 0 {
				for currRow, fValue := range childrenFieldValues {
					cValues := fValue.ChildrenList
					childrenValueRow := sheet.Row(maxRow + currRow)
					if currRow != 0 {
						for j := 0; j < currCell+1; j++ {
							if len(childrenValueRow.Cells) < j+1 {
								childrenValueRow.AddCell()
							}
						}
					}
					for _, cv := range cValues {
						if currRow > 0 {
							valueCell := childrenValueRow.AddCell()
							valueCell.Value = cv.Value
						} else {
							valueCell := fieldValueRow.AddCell()
							valueCell.Value = cv.Value
						}
					}
				}
				currCell += len(childrenFieldValues[0].ChildrenList)
			} else {
				valueCell := fieldValueRow.AddCell()
				valueCell.Value = cellValue.Value
				currCell++
				valueCell.VMerge = mergeMaxRowSize
			}
		}
	}
}
