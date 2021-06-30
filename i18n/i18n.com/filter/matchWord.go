package matchWord

import (
	"fmt"
	"regexp"
	"strings"
)

func SpliceStr(content string) (string, string, []string) {
	arr := strings.Split(content, "\\")
	var pathArr []string
	if !strings.Contains(arr[8], ".cshtml") {
		pathArr = arr[6:10]
	} else {
		pathArr = arr[6:9]
	}

	project := arr[5]
	pathStr := strings.Join(pathArr, "\\")
	sliceLength := strings.LastIndex(pathStr, ".cshtml")
	if sliceLength < 0 {
		fmt.Println("发现没有匹配到.cshtml的路径：", content, "pathStr:", pathStr)
		return "", "", make([]string, 0)
	}
	pathHandle := pathStr[:sliceLength]
	path := `Resources\` + pathHandle + "Resx"
	re := regexp.MustCompile("localizer.Localizer\\(\"([\u4e00-\u9fa5]+?.*?)\"\\)")
	matchs := re.FindAllStringSubmatch(content, 10)
	texts := make([]string, len(matchs))
	for i, mat := range matchs {
		if len(mat) > 1 {
			texts[i] = mat[1]
		}
	}

	return project, path, texts
}
