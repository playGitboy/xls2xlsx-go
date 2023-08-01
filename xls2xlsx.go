package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"strings"

	xls "github.com/shakinm/xlsReader/xls"
	"github.com/xuri/excelize/v2"
)

var existSheet1 bool
func main() {
	fmt.Print("---------------------------------------------------\n批量转换xls文件为xlsx（不依赖本机Office/WPS）\n放到xls文件所在目录，双击执行即可\n\n也可以命令行执行：\nxls2xlsx <xls文件所在目录>  \n---------------------------------------------------\n")
	var xlsfilepath string
	if len(os.Args) == 2 {
		xlsfilepath = os.Args[1]
	} else {
		xlsfilepath, _ = filepath.Abs(filepath.Dir(os.Args[0]))
	}

	xlsName := regexp.MustCompile(`(?i)(\.xls)$`)
	if xlsName.MatchString(xlsfilepath) {
		xls2xlsx(xlsfilepath)
	} else {
		os.Chdir(xlsfilepath)

		// 遍历目录
		fss := curpathfiles(xlsfilepath)
		for _, xlsfile := range *fss {
			xls2xlsx(xlsfile)
		}
	}
	fmt.Println("> 文件已全部转换完成！")
	fmt.Scanln()
}

func curpathfiles(pathname string) *[]string {
	rd, err := ioutil.ReadDir(pathname)
	if err != nil {
		fmt.Println(err)
	}

	var filenames []string
	var validID = regexp.MustCompile(`(?i)(\.xls)$`)

	for _, fi := range rd {
		if validID.MatchString(fi.Name()) {
			filenames = append(filenames, fi.Name())
		}
	}
	return &filenames
}

func xls2xlsx(xlsfile string) {
	wb1, _ := xls.OpenFile(xlsfile)
	newName := strings.TrimSuffix(xlsfile, filepath.Ext(xlsfile)) + ".xlsx"

	f2 := excelize.NewFile()
	for i := 0; i < wb1.GetNumberSheets(); i++ {
		ws1, _ := wb1.GetSheet(i)
		f2.NewSheet(ws1.GetName())
		maxrows := ws1.GetNumberRows()
		for r := 0; r < maxrows; r++ {
			r1, _ := ws1.GetRow(r)
			cols := r1.GetCols()
			for c, cs := range cols {
				ax, _ := excelize.CoordinatesToCellName(c+1, r+1)
				f2.SetCellValue(ws1.GetName(), ax, cs.GetString())
			}
		}
	}
  if len(f2.GetSheetList())>wb1.GetNumberSheets(){
    f2.DeleteSheet("Sheet1") //NewFile默认新建的Sheet1(必须填写新Sheet后才生效)
  }
	fmt.Printf("> %q 转换完成，正在保存...\n", newName)
	f2.SaveAs(newName)
	f2.Close()
}
