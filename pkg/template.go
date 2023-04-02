package pkg

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	startMatchRow int
	templates     []map[string]interface{}
)

func loadTemplate(templateName string) error {
	f, err := excelize.OpenFile(templateName)
	if err != nil {
		fmt.Println(err)
		return err
	}

	cellValue, _ := f.GetCellValue("Sheet1", "A2")
	startMatchRow, err = strconv.Atoi(cellValue)
	if err != nil {
		fmt.Println("开始匹配行数必须是数字")
		return err
	}

	row := 5
	for {
		searchColumnStr, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", row))
		if searchColumnStr == "" {
			break
		}

		searchColumns := strings.Split(searchColumnStr, " ")

		keys, _ := f.GetCellValue("Sheet1", fmt.Sprintf("B%d", row))
		if keys == "" {
			break
		}

		var targets [][]string
		col := "C"
		v := "-1"
		for v != "" {
			k, _ := f.GetCellValue("Sheet1", fmt.Sprintf("%s%d", col, row))
			v, _ = f.GetCellValue("Sheet1", fmt.Sprintf("%s%d", string(col[0]+1), row))
			if k != "" {
				targets = append(targets, []string{k, v})
				col = string(col[0] + 2)
			}
		}

		// template example:
		// {"searchColumns: ["A", "B"], "keys": ["key1", "key2"], "targets":[["k1","v1"], ["k2","v2"]}]}
		template := map[string]interface{}{
			"searchColumns": searchColumns,
			"keys":          strings.Split(keys, " "),
			"targets":       targets,
		}

		templates = append(templates, template)

		row += 1
	}
	return nil
}
