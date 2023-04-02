package pkg

import (
	"fmt"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	matchTemplateSourcePath string = "./匹配模板.xlsx"
	matchDataSourcePath     string = "./替换数据导出.xlsx"
)

func Match() {
	var err error
	defer func() {
		if err != nil {
			fmt.Printf("%s", err.Error())
			recover()
		}
	}()

	if err := loadTemplate(matchTemplateSourcePath); err != nil {
		fmt.Printf("读取模板失败, %s", err.Error())
		time.Sleep(3 * time.Second)
		return
	}

	f, err := excelize.OpenFile(matchDataSourcePath)
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
					if strings.Contains(cellValue.(string), k) {
						find = true
						break
					}
				}

				if find {
					break
				}
			}

			if find {
				for _, t := range t["targets"].([][]string) {
					if err := f.SetCellValue(sheetName, fmt.Sprintf("%s%d", t[0], row), t[1]); err != nil {
						fmt.Printf("set cell value failed, err:%+v\n", err)
					}
					// update cellMap values
					cellMap[t[0]] = t[1]
				}
			}
		}

		row += 1
	}

	if err := f.SaveAs("匹配数据导出.xlsx"); err != nil {
		fmt.Println(err)
	}
}
