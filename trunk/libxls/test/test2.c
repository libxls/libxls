/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * This file is part of libxls -- A multiplatform, C/C++ library
 * for parsing Excel(TM) files.
 *
 * Redistribution and use in source and binary forms, with or without modification, are
 * permitted provided that the following conditions are met:
 *
 *    1. Redistributions of source code must retain the above copyright notice, this list of
 *       conditions and the following disclaimer.
 *
 *    2. Redistributions in binary form must reproduce the above copyright notice, this list
 *       of conditions and the following disclaimer in the documentation and/or other materials
 *       provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY David Hoerl ''AS IS'' AND ANY EXPRESS OR IMPLIED
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
 * FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL David Hoerl OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 * SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 * ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 * Copyright 2004 Komarov Valery
 * Copyright 2006 Christophe Leitienne
 * Copyright 2008-2012 David Hoerl
 *
 */

// THIS FILE LETS YOU QUICKLY FEED A .xls FILE FOR PARSING

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "libxls/xls.h"

int main(int argc, char *argv[])
{
    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    unsigned int i;

	if(argc != 2) {
		printf("Need file arg\n");
		exit(0);
	}
	
    struct st_row_data* row;
    WORD t,tt;
    pWB=xls_open(argv[1],"UTF-8");

    if (pWB!=NULL)
    {
        for (i=0;i<pWB->sheets.count;i++)
            printf("Sheet N%i (%s) pos %i\n",i,pWB->sheets.sheet[i].name,pWB->sheets.sheet[i].filepos);

        pWS=xls_getWorkSheet(pWB,0);
        xls_parseWorkSheet(pWS);

printf("pWS->rows.lastrow %d\n", pWS->rows.lastrow);
        for (t=0;t<=pWS->rows.lastrow;t++)
        {
printf("DO IT");
            row=&pWS->rows.row[t];
            xls_showROW(row);
            for (tt=0;tt<=pWS->rows.lastcol;tt++)
            {
				xls_showCell(&row->cells.cell[tt]);
            }
        }
        printf("Count of rows: %i\n",pWS->rows.lastrow);
        printf("Max col: %i\n",pWS->rows.lastcol);
        printf("Count of sheets: %i\n",pWB->sheets.count);

        xls_showBookInfo(pWB);
    } else {
		printf("pWB == NULL\n");
	}
    return 0;
}
