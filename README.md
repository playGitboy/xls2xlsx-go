# xls2xlsx

#### 介绍
将指定目录下所有xls格式文件批量转换为xlsx格式
（常用于兼容处理用户提交文件）

优点：不依赖本机安装的Office/WPS
缺点：转换后未保留单元格格式设置

#### 说明
对gitee源码进行了重构修正：
1. xlsx库更换为通用的Excelize重构处理逻辑
2. 修复源码中扩展名识别等BUG 

#### 使用
* windows下放到xls所在目录，双击运行exe主程序
* windows下将xls目录直接拖放到exe主程序即可
* windows下拖放单/多个xls文件到exe主程序即可
* 命令行下执行 xls2xlsx <xls所在目录名>

#### 编译
1. git clone下载源码后vscode“将文件夹添加到工作区...”
2. 在vscode终端中执行"cd xls2xlsx;go mod tidy"

#### 参考
https://gitee.com/laogg/xls2xlsx  
https://xuri.me/excelize/zh-hans/workbook.html