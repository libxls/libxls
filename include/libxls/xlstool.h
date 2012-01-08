/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * This file is part of libxls -- A multiplatform, C library
 * for parsing Excel(TM) files.
 *
 * libxls is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * libxls is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with libxls.  If not, see <http://www.gnu.org/licenses/>.
 * 
 * Copyright 2004 Komarov Valery
 * Copyright 2006 Christophe Leitienne
 * Copyright 2008-2012 David Hoerl
 */

#include "xlsstruct.h"

void dumpbuf(BYTE* fname,long size,BYTE* buf);
void verbose(char* str);
BYTE* unicode_decode(const BYTE *s, int len, size_t *newlen, const char* encoding);
BYTE* get_string(BYTE *s,BYTE is2, BYTE isUnicode, char *charset);
DWORD xls_getColor(const WORD color,WORD def);

extern void xls_showBookInfo(xlsWorkBook* pWB);
extern void xls_showROW(struct st_row_data* row);
extern void xls_showColinfo(struct st_colinfo_data* col);
extern void xls_showCell(struct st_cell_data* cell);
extern void xls_showFont(struct st_font_data* font);
extern void xls_showXF(XF8* xf);
extern void xls_showFormat(struct st_format_data* format);
extern BYTE* xls_getfcell(xlsWorkBook* pWB,struct st_cell_data* cell);
extern char* xls_getCSS(xlsWorkBook* pWB);
extern void xls_showBOF(BOF* bof);
