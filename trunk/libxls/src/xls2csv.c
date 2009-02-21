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

static void OutputString(const char *string);

int main(int pintArgc, char *ptstrArgv[])
{
    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    int i;

    struct st_row_data* row;
    WORD t,tt;

	if(pintArgc != 2) {
		printf("Need file arg\n");
		exit(1);
	}

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
              printf("Feuille non trouvée");
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
                	struct st_cell_data *cell = &row->cells.cell[tt];

                    if (!cell->ishiden)
                    {
                        if (hasPreviousCol)
                          {
                           printf(";");
                          }

                        hasPreviousCol = 1;

                        // display the colspan as only one cell, but reject rowspans (they can't be converted to CSV)
                        if (cell->rowspan > 1)
                          {
                           printf("%d,%d: rowspan=%i", tt, t, cell->rowspan);
                           return 1;
                          }

                        // display the value of the cell (either numeric or string)
                        if (cell->id == 0x27e || cell->id == 0x0BD || cell->id == 0x203)
                          {
                           printf("%.15g", cell->d);
                          }
                        else if (cell->id == 0x06) // formula
                          {
                           if (cell->l == 0) // its a number
                        	 {
                        	  printf("%.15g", cell->d);
                             }
                           else
                             {
                              if (cell->str == "bool") // its boolean, and test cell->d
                                {
                            	 printf("%s", cell->d ? "true" : "false");
                                }
                              else if (cell->str == "error") // formula is in error
                                {
                                 printf("*error*");
                                }
                              else // ... cell->str is valid as the result of a string formula.
                                {
                                 OutputString(cell->str);
                                }
                             }
                          }
                        else if (cell->str!=NULL)
                          {
						   OutputString(cell->str);
                          }
                        else
                          {
                           OutputString("");
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

// Output a CSV String (between double quotes)
// Escapes (doubles) " and \ characters
static OutputString(const char *string)
{
	const char *str;

    printf("\"");
    for (str = string; *str; str++)
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
