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
程序实现代码如下：<br>

``` c
#include "tui.h"
#include "xlsx_write_interface.h"
#include "xls_read_interface.h"

void tui_xls_read_test(void)
{
	tui_xls_t * xls;
	
	xls = tui_xls_read_open("F:\\demo.xls");

	if (xls) {
		tui_xls_worksheet_t *worksheet0;
		tui_xls_worksheet_t *worksheet1;

		printf("worksheet_num = %d\n", tui_xls_read_get_worksheet_num(xls));

		worksheet0 = tui_xls_read_get_worksheet(xls, 0);
		if (worksheet0) {

			printf("worksheet0:[%d,%d]\n", tui_xls_read_get_worksheet_row_num(worksheet0), tui_xls_read_get_worksheet_column_num(worksheet0));

			printf("(0,0):%s\n", tui_xls_read_get_cell_string(worksheet0, 0, 0));
			printf("(1,0):%s\n", tui_xls_read_get_cell_string(worksheet0, 1, 0));
			printf("(2,0):%s\n", tui_xls_read_get_cell_string(worksheet0, 2, 0));
			printf("(3,0):%s\n", tui_xls_read_get_cell_string(worksheet0, 3, 0));
			printf("(4,0):%s\n", tui_xls_read_get_cell_string(worksheet0, 4, 0));
			//free worksheet0
			tui_xls_read_free_worksheet(worksheet0);
		}

		worksheet1 = tui_xls_read_get_worksheet(xls, 1);
		if (worksheet1) {

			printf("worksheet1:[%d,%d]\n", tui_xls_read_get_worksheet_row_num(worksheet1), tui_xls_read_get_worksheet_column_num(worksheet1));

			printf("(0,0):%s\n", tui_xls_read_get_cell_string(worksheet1, 0, 0));
			printf("(1,0):%s\n", tui_xls_read_get_cell_string(worksheet1, 1, 0));
			//free worksheet1
			tui_xls_read_free_worksheet(worksheet1);
		}

		//free xls
		tui_xls_read_close(xls);
	}
}
```

## `.xlsx`文件创建
下图是通过`xlsx_write_interface.h`接口，实现函数`void tui_xlsx_write_test(void)`的创建.xlsx文件。<br>
注意：在创建`demo.xlsx`文件的时会产生一个临时文件`demo.xlsx.tmp`文件，这个文件可以删除不要。
<p align="center">
<img src="https://gitee.com/tuisys/image/raw/main/xlsx_write.png">
</p>
<p align="center">
测试创建的文件
</p>

程序实现代码如下：<br>

``` c
#include "tui.h"
#include "xlsx_write_interface.h"
#include "xls_read_interface.h"

void tui_xlsx_write_test(void)
{
	tui_xlsx_t * xlsx;

	xlsx = tui_xlsx_write_open("F:\\demo.xlsx");

	if (xlsx) {
		tui_xlsx_worksheet_t * worksheet1 = tui_xlsx_add_worksheet(xlsx, "worksheet1");
		tui_xlsx_worksheet_t * worksheet2 = tui_xlsx_add_worksheet(xlsx, "worksheet2");

		tui_xlsx_format_t * format = tui_xlsx_add_format(xlsx);
		tui_xlsx_format_set_font_color(format, 0xFF0000);
		tui_xlsx_format_set_font_size(format, 50);
		tui_xlsx_format_set_align(format, XLSX_ALIGN_CENTER);
		tui_xlsx_format_set_align(format, XLSX_ALIGN_VERTICAL_CENTER);

		tui_xlsx_worksheet_set_column(worksheet1, 0, 3, 50);
		tui_xlsx_worksheet_set_row(worksheet1, 0, 60);

		/* Write some simple text. */
		tui_xlsx_worksheet_write_string(worksheet1, 0, 0, "Hello", NULL);
		tui_xlsx_worksheet_write_string(worksheet1, 1, 0, "\xe5\x93\x88\xe5\x96\xbd World", format);/* utf-8 */
		tui_xlsx_worksheet_write_number(worksheet1, 2, 0, 123, NULL);
		tui_xlsx_worksheet_write_number(worksheet1, 3, 0, 123.456, NULL);
		tui_xlsx_worksheet_write_string(worksheet1, 4, 0, "420116198503043214A", NULL);

		tui_xlsx_worksheet_write_string(worksheet2, 0, 0, "second", NULL);

		tui_xlsx_write_close(xlsx);
	}
}
```

