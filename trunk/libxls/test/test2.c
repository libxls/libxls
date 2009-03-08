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
 * Copyright 2004 Christophe Leitienne
 * Copyright 2008 David Hoerl
 */

// THIS FILE LETS YOU QUICKLY FEED A .xls FILE FOR PARSING

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include <libxls/xls.h>

int main(int argc, char *argv[])
{
    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    unsigned int i;

xls_debug=1;

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

        for (t=0;t<=pWS->rows.lastrow;t++)
        {
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
