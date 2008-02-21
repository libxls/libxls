#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <libxls/xls.h>

int main()
{

    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    FILE *f;
    int i;

    struct st_row_data* row;
    WORD t,tt;
    pWB=xls_open("files/test2.xls","KOI8-R");

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
                    if (row->cells.cell[tt].str!=NULL)
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
//    getchar();
    return 0;
}
