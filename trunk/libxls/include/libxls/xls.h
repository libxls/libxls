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
 * Copyright 2006-2009 Christophe Leitienne
 * Copyright 2008-2012 David Hoerl
 */

#ifndef XLS_INCLUDE
#define XLS_INCLUDE
 
#ifdef __cplusplus
namespace xls {
extern "C" {
#endif

#include "libxls/xlstypes.h"
#include "libxls/xlsstruct.h"
#include "libxls/xlstool.h"


#define BLANK_CELL		0x0201


extern const char* xls_getVersion(void);

extern int xls(int debug);	// Set debug. Force library to load?

extern void xls_parseWorkBook(xlsWorkBook* pWB);
extern void xls_parseWorkSheet(xlsWorkSheet* pWS);

extern xlsWorkBook* xls_open(const char *file,const char *charset);	// convert 16bit strings within the spread sheet to this 8-bit encoding (UTF-8 default)
#define xls_close xls_close_WB                  // historical
extern void xls_close_WB(xlsWorkBook* pWB);     // preferred name

extern xlsWorkSheet * xls_getWorkSheet(xlsWorkBook* pWB,int num);
extern void xls_close_WS(xlsWorkSheet* pWS);

extern xlsSummaryInfo *xls_summaryInfo(xlsWorkBook* pWB);
extern void xls_close_summaryInfo(xlsSummaryInfo *pSI);

// utility function
xlsRow *xls_row(xlsWorkSheet* pWS, WORD cellRow);
xlsCell	*xls_cell(xlsWorkSheet* pWS, WORD cellRow, WORD cellCol);

#ifdef __cplusplus
} // extern c block
} // namespace
#endif

#endif

