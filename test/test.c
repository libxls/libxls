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

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include <libxls/xls.h>

int main(int argc, char *argv[])
{

    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    FILE *f;
    unsigned int i;

    struct st_row_data* row;
    WORD t,tt;
    pWB=xls_open("files/test2.xls", "ASCII"); // "KOI8-R"

    if (pWB!=NULL)
    {
        f=fopen ("test.htm", "w");
        for (i=0;i<pWB->sheets.count;i++)
            printf("Sheet N%i (%s) pos %i\n",i,pWB->sheets.sheet[i].name,pWB->sheets.sheet[i].filepos);

        pWS=xls_getWorkSheet(pWB,0);
        xls_parseWorkSheet(pWS);
        fprintf(f,"<style type=\"text/css\">\n%s</style>\n",xls_getCSS(pWB));
        fprintf(f,"<table border=0 cellspacing=0 cellpadding=2>");

        for (t=0;t<=pWS->rows.lastrow;t++)
        {
            row=&pWS->rows.row[t];
            //		xls_showROW(row->row);
            fprintf(f,"<tr>");
            for (tt=0;tt<=pWS->rows.lastcol;tt++)
            {
                if (!row->cells.cell[tt].ishiden)
                {
                    fprintf(f,"<td");
                    if (row->cells.cell[tt].colspan)
                        fprintf(f," colspan=%i",row->cells.cell[tt].colspan);
                    //				if (t==0) fprintf(f," width=%i",row->cells.cell[tt].width/35);
                    if (row->cells.cell[tt].rowspan)
                        fprintf(f," rowspan=%i",row->cells.cell[tt].rowspan);
                    fprintf(f," class=xf%i",row->cells.cell[tt].xf);
                    fprintf(f,">");
                    if (row->cells.cell[tt].str!=NULL && row->cells.cell[tt].str[0]!='\0')
                        fprintf(f,"%s",row->cells.cell[tt].str);
                    else
                        fprintf(f,"%s","&nbsp;");
                    fprintf(f,"</td>");
                }
            }
            fprintf(f,"</tr>\n");
        }
        fprintf(f,"</table>");
        printf("Count of rows: %i\n",pWS->rows.lastrow);
        printf("Max col: %i\n",pWS->rows.lastcol);
        printf("Count of sheets: %i\n",pWB->sheets.count);

        fclose(f);
        xls_showBookInfo(pWB);
    }

    return 0;
}
