package pkg

import (
	"fmt"
	"regexp"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	replaceTemplateSourcePath string = "./替换模板.xlsx"
	replaceDataSourcePath     string = "./原始数据.xlsx"
)

func Replace() {
	var err error
	defer func() {
		if err != nil {
			fmt.Printf("%s", err.Error())
			recover()
		}
	}()

	err = loadTemplate(replaceTemplateSourcePath)
	if err != nil {
		fmt.Println(err)
		time.Sleep(3 * time.Second)
		return
	}

	f, err := excelize.OpenFile(replaceDataSourcePath)
	if err != nil {
		fmt.Println(err)
		return
	}

	sheetName := f.GetSheetName(f.GetActiveSheetIndex())
	rows, err := f.Rows(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}

	row := startMatchRow
	for rows.Next() {
		fmt.Printf("第%d行\n", row)

		cellMap := map[string]interface{}{}

		for _, t := range templates {
			find := false
			matchCellValue := ""
			replaceCellValue := ""
			var cellValues []interface{}

			for _, n := range t["searchColumns"].([]string) {

				if _, ok := cellMap[n]; !ok {
					cellMap[n], _ = f.GetCellValue(sheetName, fmt.Sprintf("%s%d", n, row))
				}

				cellValues = append(cellValues, cellMap[n])
			}

			for _, cellValue := range cellValues {
				if cellValue == "" {
					continue
				}
				for _, k := range t["keys"].([]string) {
					if r := reFirstMatch(cellValue.(string), k); r != "" {
						find = true
						matchCellValue = cellValue.(string)
						replaceCellValue = r
						break
					}
				}

				if find {
					break
				}
			}

			if find {
				// 替换
				for _, t := range t["targets"].([][]string) {
					if err := f.SetCellValue(sheetName, fmt.Sprintf("%s%d", t[0], row), strings.ReplaceAll(matchCellValue, replaceCellValue, t[1])); err != nil {
						fmt.Printf("set cell value failed, err:%+v\n", err)
					}
				}
			}
		}

		row += 1
	}

	if err := f.SaveAs("替换数据导出.xlsx"); err != nil {
		fmt.Println(err)
	}
}

func reFirstMatch(str, substr string) string {
	substr = strings.ReplaceAll(substr, "*", ".*")
	substr = strings.ReplaceAll(substr, "(", "\\(")
	substr = strings.ReplaceAll(substr, ")", "\\)")
	r, err := regexp.Compile(substr)
	if err != nil {
		fmt.Println(err.Error())
	}
	return r.FindString(str)
}
