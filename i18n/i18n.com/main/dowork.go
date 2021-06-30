package main

import (
	"encoding/json"
	"fmt"
	"os"
	"sort"
	"strings"
	"time"

	"i18n.com/matchWord"

	"github.com/tealeg/xlsx"
)

func main() {
	//生成中间excel文件
	originalPath := "../system0630.xlsx"
	systemOriginals := ProductMiddleExcel(originalPath, "")

	//导出的文件添加没有的项
	exportPath := "../translate0630.xlsx"
	exportExcel := ReadExportExcel(exportPath)
	//获取没有翻译的项
	waittingHandle := make([]MiddleExcel, 0)
	flag := false
	for _, item := range systemOriginals {
		for _, e := range exportExcel {
			if item.ResxPath == e.ResxPath && item.Texts == e.Texts {
				flag = true
				break
			}
		}
		if !flag {
			waittingHandle = append(waittingHandle, item)
		}
		flag = false
	}

	sumTranslateTexts := ReadTranslateText("../sum_translate.xlsx")
	for i, waitItem := range waittingHandle {
		for _, sumItem := range sumTranslateTexts {
			if waitItem.Texts == sumItem.Name {
				waittingHandle[i].TextEn = sumItem.Value
			}
		}
	}
	// 把新增项写入导出的文件中 直接导入该文件即可
	// WriteExportExcel(exportPath, &waittingHandle)
	waitmap := splitSlice(waittingHandle)
	// var str string
	// 	<data name="取消" xml:space="preserve">
	//     <value>取消</value>
	//   </data>
	var path string
	for _, m := range waitmap {
		if len(m) <= 0 {
			continue
		}
		path = "../" + strings.Replace(m[0].ResxPath, "\\", "/", -1)
		err := os.MkdirAll(path, os.ModePerm)
		if err != nil {
			panic(err)
		}
		sxml, err := os.Create(path + "/sxml.xml")
		if err != nil {
			panic(err)
		}
		var str string
		for _, item := range waittingHandle {
			str += fmt.Sprintf(`<data name="%s" xml:space="preserve">`, item.Texts)
			//生成英文xml
			if item.TextEn == "" {
				str += fmt.Sprintf(`<value>%s</value></data>`, item.Texts)
			} else {
				str += fmt.Sprintf(`<value>%s</value></data>`, item.TextEn)
			}
			//中文xml
			str += fmt.Sprintf(`<value>%s</value></data>`, item.Texts)
		}

		defer sxml.Close()
		sxml.WriteString(str)
	}

	// //读取翻译文本到map 并生成json
	// t210 := "../t_210624.xlsx"
	// res := ReadExport210(t210)
	// jsonObj, _ := json.Marshal(res)
	// str := string(jsonObj)
	// sjson, err := os.Create("../sjson.json")
	// if err != nil {
	// 	panic(err)
	// }
	// defer sjson.Close()
	// sjson.WriteString(str)

}

type ProductExcel struct {
	ProjectPath string
	ResxPath    string
	Texts       []string
}

type MiddleExcel struct {
	ProjectPath string
	ResxPath    string
	Texts       string
	TextEn      string
}

type MiddleExcelDecrement []MiddleExcel

func (s MiddleExcelDecrement) Len() int { return len(s) }

func (s MiddleExcelDecrement) Swap(i, j int) { s[i], s[j] = s[j], s[i] }

func (s MiddleExcelDecrement) Less(i, j int) bool {
	return s[i].ResxPath > s[j].ResxPath
}

func ReadExportExcel(exportPath string) []MiddleExcel {
	exportSheetName := "ResXResourceManager"
	wb, err := xlsx.OpenFile(exportPath)
	if err != nil {
		panic(err)
	}
	fmt.Println("--translate0630.xlsx中包含的sheet--")
	for i, sh := range wb.Sheets {
		fmt.Println(i, sh.Name)
	}
	fmt.Println("--translate0630.xlsx--")
	//获取ResXResourceManager
	exportSheet, ok := wb.Sheet[exportSheetName]
	if !ok {
		panic("Sheet does not exist")
	}
	fmt.Println("Max row in ResXResourceManager Sheet:", exportSheet.MaxRow)
	var row *xlsx.Row
	middles := make([]MiddleExcel, exportSheet.MaxRow)
	for i := 0; i < exportSheet.MaxRow; i++ {
		row = exportSheet.Row(i)
		if err != nil {
			panic(err)
		}
		tmpMiddle := MiddleExcel{
			ProjectPath: row.Cells[0].Value,
			ResxPath:    row.Cells[1].Value,
			Texts:       row.Cells[2].Value,
		}
		middles[i] = tmpMiddle
	}
	return middles
}

func WriteExportExcel(exportPath string, waittings *[]MiddleExcel) {
	var wb *xlsx.File
	var sh *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell
	var err error
	var ok bool
	wb, err = xlsx.OpenFile(exportPath)
	if err != nil {
		fmt.Println("打开文件", exportPath, "失败")
		panic(err)
	}
	sh, ok = wb.Sheet["ResXResourceManager"]
	if !ok {
		panic("在读取导入文件时未获取到sheet")
	}
	for _, item := range *waittings {
		row = sh.AddRow()
		cell = row.AddCell()
		cell.Value = item.ProjectPath
		cell = row.AddCell()
		cell.Value = item.ResxPath
		cell = row.AddCell()
		cell.Value = item.Texts
		cell = row.AddCell()
		cell.Value = ""
		cell = row.AddCell()
		cell.Value = item.Texts
		cell = row.AddCell()
		cell.Value = ""
		cell = row.AddCell()
		cell.Value = item.TextEn
	}
	err = wb.Save(exportPath)
	if err != nil {
		panic(err)
	}
}

func ProductMiddleExcel(originalPath string, productPath string) []MiddleExcel {
	if originalPath == "" {
		panic("请输入原始excel文件路径")
	}
	if productPath == "" {
		now := time.Now()
		nowStr := fmt.Sprintf("%d-%d-%d-%d-%d-%d", now.Year(), now.Month(), now.Day(), now.Hour(), now.Minute(), now.Second())
		productPath = `../middleExcel_` + nowStr + `.xlsx`
	}
	fmt.Println("开始读取", originalPath, "……")
	wb, err := xlsx.OpenFile(originalPath)
	if err != nil {
		panic(err)
	}
	sh, ok := wb.Sheet["Sheet1"]
	if !ok {
		panic("未获取到sheet")
	}
	productRows := make([]ProductExcel, sh.MaxRow)
	fmt.Println("开始匹配路径和翻译文本……")
	for i := 0; i < sh.MaxRow; i++ {
		row := sh.Row(i)
		content := row.Cells[0].Value
		projectPath, resxPath, texts := matchWord.SpliceStr(content)
		productRows[i].ProjectPath = projectPath
		productRows[i].ResxPath = resxPath
		productRows[i].Texts = texts
	}
	middleRows := make([]MiddleExcel, 0)
	for _, i := range productRows {
		for _, j := range i.Texts {
			tmpMiddle := MiddleExcel{
				ProjectPath: i.ProjectPath,
				ResxPath:    i.ResxPath,
				Texts:       j,
			}
			middleRows = append(middleRows, tmpMiddle)
		}

	}
	//格式为E:\Source\Repo\fcm_tmcWeb\src\Ontheway.TMC.AppViews.DragonView\Views\ApplyApprove\Pass.cshtml(24):@localizer.Localizer("您正在审批") <strong>@localizer.Localizer("申请单号")：@Model.SerialNumber</strong>
	if len(middleRows) > 0 {
		//二次匹配

		middleRows = DeDuplicatesAndEmpty(middleRows)
		fmt.Println("文件:", originalPath, "共", len(middleRows), "行")
		fmt.Println("开始生成" + productPath + "……")
		var productWb *xlsx.File
		var middleSh *xlsx.Sheet
		var row *xlsx.Row
		var cell *xlsx.Cell
		var err error
		productWb = xlsx.NewFile()
		middleSh, err = productWb.AddSheet("middleSheet")
		if err != nil {
			fmt.Println("创建middleSheet异常……")
			panic(err)
		}

		for _, n := range middleRows {
			row = middleSh.AddRow()
			cell = row.AddCell()
			cell.Value = n.ProjectPath
			cell = row.AddCell()
			cell.Value = n.ResxPath
			cell = row.AddCell()
			cell.Value = n.Texts
		}
		err = productWb.Save(productPath)
		if err != nil {
			panic(err)
		}
		fmt.Println("中间文件生成完成……")
	}
	return middleRows
}

func DeDuplicatesAndEmpty(a []MiddleExcel) (ret []MiddleExcel) {
	sort.Stable(MiddleExcelDecrement(a))
	a_len := len(a)
	flag := false
	for i := 0; i < a_len; i++ {
		for j := 0; j < len(ret); j++ {
			if a[i].Texts == ret[j].Texts && a[i].ResxPath == ret[j].ResxPath {
				flag = true
				break
			}
		}
		if !flag {
			ret = append(ret, a[i])
		}
		flag = false
	}
	return
}

type KeyValue struct {
	Name  string
	Value string
}

func ReadExport210(exportPath string) []KeyValue {
	exportSheetName := "Tickets"
	wb, err := xlsx.OpenFile(exportPath)
	if err != nil {
		panic(err)
	}
	fmt.Println("--t_210624.xlsx中包含的sheet--")
	for i, sh := range wb.Sheets {
		fmt.Println(i, sh.Name)
	}
	fmt.Println("--t_210624.xlsx--")
	//获取ResXResourceManager
	exportSheet, ok := wb.Sheet[exportSheetName]
	if !ok {
		panic("Sheet does not exist")
	}
	fmt.Println("Max row in ResXResourceManager Sheet:", exportSheet.MaxRow)
	var row *xlsx.Row
	res := make([]KeyValue, 0)
	for i := 0; i < exportSheet.MaxRow; i++ {
		row = exportSheet.Row(i)
		if err != nil {
			panic(err)
		}
		if row.Cells[3].Value != "" {
			tmp := KeyValue{
				Name:  row.Cells[2].Value,
				Value: row.Cells[2].Value,
			}
			res = append(res, tmp)
		}
	}
	return res
}

func ReadTranslateText(exportPath string) []KeyValue {
	var exportSheet *xlsx.Sheet
	var row *xlsx.Row
	var ok bool
	res := make([]KeyValue, 0)
	wb, err := xlsx.OpenFile(exportPath)
	if err != nil {
		panic(err)
	}
	exportSheet, ok = wb.Sheet["Sheet1"]
	if !ok {
		panic("Sheet does not exist")
	}
	fmt.Println("Max row in Sheet1 Sheet:", exportSheet.MaxRow)
	for i := 0; i < exportSheet.MaxRow; i++ {
		row = exportSheet.Row(i)
		if err != nil {
			panic(err)
		}
		if len(row.Cells) > 2 && row.Cells[2].Value != "" {
			tmp := KeyValue{
				Name:  row.Cells[1].Value,
				Value: row.Cells[2].Value,
			}
			res = append(res, tmp)
		}
	}

	exportSheet, ok = wb.Sheet["Sheet2"]
	if !ok {
		panic("Sheet does not exist")
	}
	fmt.Println("Max row in Sheet2 Sheet:", exportSheet.MaxRow)
	for i := 0; i < exportSheet.MaxRow; i++ {
		row = exportSheet.Row(i)
		if err != nil {
			panic(err)
		}
		if len(row.Cells) > 2 && row.Cells[2].Value != "" {
			tmp := KeyValue{
				Name:  row.Cells[1].Value,
				Value: row.Cells[2].Value,
			}
			res = append(res, tmp)
		}
	}

	exportSheet, ok = wb.Sheet["Sheet3"]
	if !ok {
		panic("Sheet does not exist")
	}
	fmt.Println("Max row in Sheet3 Sheet:", exportSheet.MaxRow)
	for i := 0; i < exportSheet.MaxRow; i++ {
		row = exportSheet.Row(i)
		if err != nil {
			panic(err)
		}
		if len(row.Cells) > 3 && row.Cells[3].Value != "" {
			tmp := KeyValue{
				Name:  row.Cells[2].Value,
				Value: row.Cells[3].Value,
			}
			res = append(res, tmp)
		}
	}

	exportSheet, ok = wb.Sheet["Sheet4"]
	if !ok {
		panic("Sheet does not exist")
	}
	fmt.Println("Max row in Sheet4 Sheet:", exportSheet.MaxRow)
	for i := 0; i < exportSheet.MaxRow; i++ {
		row = exportSheet.Row(i)
		if err != nil {
			panic(err)
		}
		if len(row.Cells) > 3 && row.Cells[3].Value != "" {
			tmp := KeyValue{
				Name:  row.Cells[2].Value,
				Value: row.Cells[3].Value,
			}
			res = append(res, tmp)
		}
	}

	exportSheet, ok = wb.Sheet["Sheet5"]
	if !ok {
		panic("Sheet does not exist")
	}
	fmt.Println("Max row in Sheet5 Sheet:", exportSheet.MaxRow)
	for i := 0; i < exportSheet.MaxRow; i++ {
		row = exportSheet.Row(i)
		if err != nil {
			panic(err)
		}
		if len(row.Cells) > 6 && row.Cells[6].Value != "" {
			tmp := KeyValue{
				Name:  row.Cells[4].Value,
				Value: row.Cells[6].Value,
			}
			res = append(res, tmp)
		}
	}
	tjson, err := os.Create("../sum_translate.json")
	if err != nil {
		panic(err)
	}
	tjsonObj1, err := json.Marshal(res)
	if err != nil {
		panic(err)
	}
	tjsonStr1 := string(tjsonObj1)
	tjson.WriteString(tjsonStr1)
	defer tjson.Close()
	return res
}

func splitSlice(list []MiddleExcel) [][]MiddleExcel {
	sort.Sort(MiddleExcelDecrement(list))
	returnData := make([][]MiddleExcel, 0)
	i := 0
	var j int
	for {
		if i >= len(list) {
			break
		}
		for j = i + 1; j < len(list) && list[i].ResxPath == list[j].ResxPath; j++ {
		}

		returnData = append(returnData, list[i:j])
		i = j
	}
	return returnData
}
