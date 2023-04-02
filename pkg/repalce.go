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
			// 是否匹配
			find := false
			// 匹配的值
			matchCellValue := ""
			// 需要替换的部分值
			replaceCellValue := ""

			// 缓存匹配到的单元格的值
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
					newValue := strings.ReplaceAll(matchCellValue, replaceCellValue, t[1])
					if err := f.SetCellValue(sheetName, fmt.Sprintf("%s%d", t[0], row), newValue); err != nil {
						fmt.Printf("set cell value failed, err:%+v\n", err)
					}

					// update cellMap values
					cellMap[t[0]] = newValue
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
