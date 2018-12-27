package main

import (
	"github.com/tealeg/xlsx"
	"strconv"
	"zhaobin.com/report/model"
	"zhaobin.com/report/xlsxExport"
)

func main() {
	var sheetName = "sheetName"
	var titleName = "titleName"
	var fieldKeys []*model.FieldKey
	for i := 0; i < 10; i++ {
		fieldKey := &model.FieldKey{Id: "kid" + strconv.Itoa(i), Name: "k" + strconv.Itoa(i)}

		if i == 3 || i == 5 {
			var childers []*model.FieldKey
			for x := 0; x < 4; x++ {
				childers = append(childers, &model.FieldKey{Id: "cid" + strconv.Itoa(x), Name: "ck" + strconv.Itoa(x)})
			}
			fieldKey.ChildrenList = childers
		}

		fieldKeys = append(fieldKeys, fieldKey)
	}

	var fieldValues []*model.FieldValue
	for j := 0; j < 2; j++ {
		fieldValue := &model.FieldValue{Value: "v" + strconv.Itoa(j)}
		var childrenValues []*model.FieldValue
		for i := 0; i < 10; i++ {
			rowValue := &model.FieldValue{Value: "vi" + strconv.Itoa(i)}

			if i == 3 || i == 5 {
				var fChildrens []*model.FieldValue
				for k := 0; k < 3; k++ {
					fValue := &model.FieldValue{Value: "jii" + strconv.Itoa(i)}
					var childers []*model.FieldValue
					for x := 0; x < 4; x++ {
						childers = append(childers, &model.FieldValue{Value: "x" + strconv.Itoa(x)})
					}
					fValue.ChildrenList = childers
					fChildrens = append(fChildrens, fValue)
				}
				rowValue.ChildrenList = fChildrens
			}

			childrenValues = append(childrenValues, rowValue)
		}
		fieldValue.ChildrenList = childrenValues
		fieldValues = append(fieldValues, fieldValue)
	}
	export := &xlsxExport.Export{TitleName: titleName, SheetName: sheetName, FieldKeys: fieldKeys, FieldValues: fieldValues}
	export.File = "gender2.xlsx"

	style := xlsx.NewStyle()
	font := xlsx.NewFont(20, "Verdana")
	// font.Color = "3733DC"
	// f := &xlsx.Font{Color: "#3733DC", Size: 20}
	style.Font = *font
	fill := xlsx.NewFill("solid", "3733DC", "3733DC")
	style.Fill = *fill

	alig := &xlsx.Alignment{
		Horizontal: "center",
		Vertical:   "center",
	}
	style.Alignment = *alig
	export.TitleStyle = style
	export.KeyFirstRowStyle = style
	export.Save()
}
