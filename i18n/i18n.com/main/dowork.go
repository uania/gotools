package main

import (
	"fmt"
	"sort"
	"time"

	"i18n.com/matchWord"

	"github.com/tealeg/xlsx"
)

func main() {
	//生成中间excel文件
	originalPath := "../system0629.xlsx"
	ProductMiddleExcel(originalPath, "")

	//导出的文件添加没有的项
	// exportPath := "../translate0629.xlsx"
	// ReadExportExcel(exportPath)
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
}

type MiddleExcelDecrement []MiddleExcel

func (s MiddleExcelDecrement) Len() int { return len(s) }

func (s MiddleExcelDecrement) Swap(i, j int) { s[i], s[j] = s[j], s[i] }

func (s MiddleExcelDecrement) Less(i, j int) bool {
	return s[i].Texts > s[j].Texts && s[i].ResxPath > s[j].ResxPath
}

func ReadExportExcel(exportPath string) []MiddleExcel {
	exportSheetName := "ResXResourceManager"
	wb, err := xlsx.OpenFile(exportPath)
	if err != nil {
		panic(err)
	}
	fmt.Println("--translate0629.xlsx中包含的sheet--")
	for i, sh := range wb.Sheets {
		fmt.Println(i, sh.Name)
	}
	fmt.Println("--translate0629.xlsx--")
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

func ProductMiddleExcel(originalPath string, productPath string) {
	if originalPath == "" {
		panic("请输入原始excel文件路径")
	}
	if productPath == "" {
		now := time.Now()
		nowStr := fmt.Sprintf("%d-%d-%d", now.Hour(), now.Minute(), now.Second())
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
}

func DeDuplicatesAndEmpty(a []MiddleExcel) (ret []MiddleExcel) {
	sort.Stable(MiddleExcelDecrement(a))
	a_len := len(a)
	for i := 0; i < a_len; i++ {
		if (i > 0 && a[i-1].Texts == a[i].Texts && a[i-1].ResxPath == a[i].ResxPath) || len(a[i].Texts) == 0 {
			continue
		}
		ret = append(ret, a[i])
	}
	return
}
