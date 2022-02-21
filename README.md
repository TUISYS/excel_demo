<h1 align="center"> TUI Excel表格测试教程 </h1>

## 描述
TUI V2.1以上版本支持对`.xls`文件读取，和`.xlsx`文件创建。都是UTF-8编码字符串的输出和输入。

## `.xls`文件读取
下图是要读取的xls文件
<p align="center">
<img src="https://gitee.com/tuisys/image/raw/main/xls_read.png">
</p>
<p align="center">
需要读取的.xls文件
</p>

下图是通过`xls_read_interface.h`接口，实现函数`void tui_xls_read_test(void)`的`UTF-8`输出
<p align="center">
<img src="https://gitee.com/tuisys/image/raw/main/xls_read_output.png">
</p>
<p align="center">
测试输出的打印
</p>

## `.xlsx`文件创建
下图是通过`xlsx_write_interface.h`接口，实现函数`void tui_xlsx_write_test(void)`的创建.xlsx文件
<p align="center">
<img src="https://gitee.com/tuisys/image/raw/main/xlsx_write.png">
</p>
<p align="center">
测试创建的文件
</p>
