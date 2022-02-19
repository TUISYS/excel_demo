#ifndef __XLSX_WRITE_INTERFACE_H__
#define __XLSX_WRITE_INTERFACE_H__

#ifndef TUI_SOURCE_CODE
#include "stdxxx.h"
#else
#include "def_inlcude.h"
#endif

typedef void * tui_xlsx_t;
typedef void * tui_xlsx_worksheet_t;
typedef void * tui_xlsx_format_t;

/** Alignment values for tui_xlsx_format_set_align(). */
typedef enum tag_tui_xlsx_format_align {
	XLSX_ALIGN_NONE = 0,

	/** Left horizontal alignment */
	XLSX_ALIGN_LEFT,

	/** Center horizontal alignment */
	XLSX_ALIGN_CENTER,

	/** Right horizontal alignment */
	XLSX_ALIGN_RIGHT,

	/** Cell fill horizontal alignment */
	XLSX_ALIGN_FILL,

	/** Justify horizontal alignment */
	XLSX_ALIGN_JUSTIFY,

	/** Center Across horizontal alignment */
	XLSX_ALIGN_CENTER_ACROSS,

	/** Left horizontal alignment */
	XLSX_ALIGN_DISTRIBUTED,

	/** Top vertical alignment */
	XLSX_ALIGN_VERTICAL_TOP,

	/** Bottom vertical alignment */
	XLSX_ALIGN_VERTICAL_BOTTOM,

	/** Center vertical alignment */
	XLSX_ALIGN_VERTICAL_CENTER,

	/** Justify vertical alignment */
	XLSX_ALIGN_VERTICAL_JUSTIFY,

	/** Distributed vertical alignment */
	XLSX_ALIGN_VERTICAL_DISTRIBUTED
} tui_xlsx_format_align_e;

tui_xlsx_t * tui_xlsx_write_open(const char *path);
void tui_xlsx_write_close(tui_xlsx_t * xls);
tui_xlsx_worksheet_t * tui_xlsx_add_worksheet(tui_xlsx_t * xls, const char * name);
tui_xlsx_format_t * tui_xlsx_add_format(tui_xlsx_t * xls);
void tui_xlsx_format_set_font_color(tui_xlsx_format_t *format, uint32_t color);
void tui_xlsx_format_set_font_size(tui_xlsx_format_t *format, uint16_t size);
void tui_xlsx_format_set_align(tui_xlsx_format_t *format, tui_xlsx_format_align_e align);
int tui_xlsx_worksheet_set_column(tui_xlsx_worksheet_t *worksheet, uint16_t firstcol, uint16_t lastcol, uint32_t width);
int tui_xlsx_worksheet_set_row(tui_xlsx_worksheet_t *worksheet, uint32_t row_num, int height);
int tui_xlsx_worksheet_write_string(tui_xlsx_worksheet_t *worksheet, uint32_t row_num, uint16_t col_num, const char *string, tui_xlsx_format_t *format);
int tui_xlsx_worksheet_write_number(tui_xlsx_worksheet_t *worksheet, uint32_t row_num, uint16_t col_num, double value, tui_xlsx_format_t *format);


#endif /* __XLSX_WRITE_INTERFACE_H__ */

