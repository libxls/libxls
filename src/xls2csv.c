#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <xls.h>

int main(int pintArgc, char *ptstrArgv[])
{
    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    int i;

    struct st_row_data* row;
    WORD t,tt;

    // open workbook, choose standard conversion
    pWB=xls_open(ptstrArgv[1], "iso-8859-15//TRANSLIT");

    // process workbook if found
    if (pWB!=NULL)
    {
        // check if the requested sheet (if any) exists
        if (  (pintArgc >= 3)
            &&(strcmp(ptstrArgv[2], "-l") != 0) )
          {
           for (i=0;i<pWB->sheets.count;i++)
              {
               if (strcmp(ptstrArgv[2], pWB->sheets.sheet[i].name) == 0)
                 {
                  break;
                 }
              }

           if (i == pWB->sheets.count)
             {
              printf("Feuille non trouv‚e");
              return EXIT_FAILURE;
             }
          }

        // process all sheets
        for (i=0;i<pWB->sheets.count;i++)
           {
            int lineWritten = 0;

            // check if this is a requested sheet
            if (pintArgc >= 3)
              {
               if (strcmp(ptstrArgv[2], "-l") == 0)
                 {
                  printf("%s\n", pWB->sheets.sheet[i].name);
                  continue;
                 }
               if (strcmp(ptstrArgv[2], pWB->sheets.sheet[i].name) != 0)
                 {
                  continue;
                 }
              }

            // open and parse the sheet
            pWS=xls_getWorkSheet(pWB,i);
            xls_parseWorkSheet(pWS);

            // process all rows of the sheet
            for (t=0;t<=pWS->rows.lastrow;t++)
            {
                int hasPreviousCol = 0;
                row=&pWS->rows.row[t];

                // process cells
                if (lineWritten)
                  {
                   printf("\n");
                  }
                else
                  {
                   lineWritten = 1;
                  }

                for (tt=0;tt<=pWS->rows.lastcol;tt++)
                {
                    if (!row->cells.cell[tt].ishiden)
                    {
                        if (hasPreviousCol)
                          {
                           printf(";");
                          }

                        hasPreviousCol = 1;

                        // display the colspan as only one cell, but reject rowspans (they can't be converted to CSV)
                        if (row->cells.cell[tt].rowspan > 1)
                          {
                           printf("%d,%d: rowspan=%i", tt, t, row->cells.cell[tt].rowspan);
                           return 1;
                          }

                        // display the value of the cell (either numeric or string)
                        if (row->cells.cell[tt].id==0x27e || row->cells.cell[tt].id==0x0BD || row->cells.cell[tt].id==0x203)
                          {
                           printf("%.15g", row->cells.cell[tt].d);
                          }
                        else if (row->cells.cell[tt].str!=NULL)
                          {
                           char *str = row->cells.cell[tt].str;

                           printf("\"");
                           for (str = row->cells.cell[tt].str; *str; str++)
                              {
                               if (*str == '\"')
                                 {
                                  printf("\"\"");
                                 }
                               else if (*str == '\\')
                                 {
                                  printf("\\\\");
                                 }
                               else
                                 {
                                  printf("%c", *str);
                                 }
                              }
                           printf("\"");
                          }
                        else
                          {
                           printf("\"\"");
                          }
                    }
                }
            }
           }

         xls_close(pWB);
         return EXIT_SUCCESS;
    }
    
    return EXIT_FAILURE;
}


