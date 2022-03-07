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

int main(void)
{
	tui_xls_read_test();
	tui_xlsx_write_test();

}

