#ifndef __XLS_READ_INTERFACE_H__
#define __XLS_READ_INTERFACE_H__

#ifndef TUI_SOURCE_CODE
#include "stdxxx.h"
#else
#include "def_inlcude.h"
#endif

typedef void * tui_xls_t;
typedef void * tui_xls_worksheet_t;


tui_xls_t * tui_xls_read_open(const char *path);
void tui_xls_read_close(tui_xls_t * xls);

int tui_xls_read_get_worksheet_num(tui_xls_t *xls);

tui_xls_worksheet_t * tui_xls_read_get_worksheet(tui_xls_t *xls, unsigned int worksheet_index);
void tui_xls_read_free_worksheet(tui_xls_worksheet_t *worksheet);

int tui_xls_read_get_worksheet_row_num(tui_xls_worksheet_t *worksheet);
int tui_xls_read_get_worksheet_column_num(tui_xls_worksheet_t *worksheet);
const char* tui_xls_read_get_cell_string(tui_xls_worksheet_t *worksheet, int row_index, int column_index);



#endif /* __XLS_READ_INTERFACE_H__ */

