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

int main(int argc, char *argv[])
{
    xlsWorkBook* pWB;
    xlsWorkSheet* pWS;
    unsigned int i;

    struct st_row_data* row;
    WORD cellRow,cellCol;

	if(argc < 2) {
		printf("Need file arg\n");
		exit(1);
	}

    // open workbook, choose standard conversion
    pWB=xls_open(argv[1], "iso-8859-15//TRANSLIT");

    // process workbook if found
    if (pWB!=NULL)
    {
        // check if the requested sheet (if any) exists
        if (  (argc >= 3)
            &&(strcmp(argv[2], "-l") != 0) )
          {
           for (i=0;i<pWB->sheets.count;i++)
              {
               if (strcmp(argv[2], pWB->sheets.sheet[i].name) == 0)
                 {
                  break;
                 }
              }

           if (i == pWB->sheets.count)
             {
              printf("Sheet not found");
              return EXIT_FAILURE;
             }
          }

        // process all sheets
        for (i=0;i<pWB->sheets.count;i++)
           {
            int lineWritten = 0;

            // check if this is a requested sheet
            if (argc >= 3)
              {
               if (strcmp(argv[2], "-l") == 0)
                 {
                  printf("%s\n", pWB->sheets.sheet[i].name);
                  continue;
                 }
               if (strcmp(argv[2], pWB->sheets.sheet[i].name) != 0)
                 {
                  continue;
                 }
              }

            // open and parse the sheet
            pWS=xls_getWorkSheet(pWB,i);
            xls_parseWorkSheet(pWS);

            // process all rows of the sheet
            for (cellRow=0;cellRow<=pWS->rows.lastrow;cellRow++)
            {
                int hasPreviousCol = 0;
                row = xls_row(pWS, cellRow);

                // process cells
                if (lineWritten)
                  {
                   printf("\n");
                  }
                else
                  {
                   lineWritten = 1;
                  }

                for (cellCol=0;cellCol<=pWS->rows.lastcol;cellCol++)
                {
                	xlsCell	*cell = xls_cell(pWS, cellRow, cellCol);

                    if (  (cell)
                    	&&(!cell->ishiden) )
                    {
                        if (hasPreviousCol)
                          {
                           printf(";");
                          }

                        hasPreviousCol = 1;

                        // display the colspan as only one cell, but reject rowspans (they can't be converted to CSV)
                        if (cell->rowspan > 1)
                          {
                           printf("%d,%d: rowspan=%i", cellCol, cellRow, cell->rowspan);
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
                            	 printf("%s", (int)cell->d ? "true" : "false");
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
static void OutputString(const char *string)
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
