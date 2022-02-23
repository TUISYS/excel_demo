<h1 align="center"> TUI Excel表格测试教程 </h1>

## 描述
TUI *V2.1* 以上版本支持对`.xls`文件读取，和`.xlsx`文件创建，都是UTF-8编码字符串的输出和输入。<br>
`TUISYS/excel_demo`仓库只提供了头文件接口和测试demo，实际应用要结合`TUISYS/tui_project`仓库里面的tui库文件。
直接将`xls_read_interface.h`和`xlsx_write_interface.h`放进`TUISYS/tui_project/includes`，然后应用就可以调用相关函数了。

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
下图是通过`xlsx_write_interface.h`接口，实现函数`void tui_xlsx_write_test(void)`的创建.xlsx文件。<br>
注意：在创建`demo.xlsx`文件的时会产生一个临时文件`demo.xlsx.tmp`文件，这个文件可以删除不要。
<p align="center">
<img src="https://gitee.com/tuisys/image/raw/main/xlsx_write.png">
</p>
<p align="center">
测试创建的文件
</p>
