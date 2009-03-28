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
 * Copyright 2004-2009 Christophe Leitienne
 * Copyright 2008 David Hoerl
 */

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include <libxls/xls.h>

const char stringSeparator = '\"';
const char *lineSeparator = "\n";
const char *fieldSeparator = ";";

static void OutputString(const char *string);
static void OutputNumber(const double number);

int main(int argc, char *argv[]) {
	xlsWorkBook* pWB;
	xlsWorkSheet* pWS;
	unsigned int i;

	struct st_row_data* row;
	WORD cellRow, cellCol;

	// check argument count
	if (argc < 2 || argc > 3) {
		fprintf(stderr, "usage: %s <Excel file> [-l|<sheet>]\n", argv[0]);
		fprintf(stderr, "       display cells of an Excel file as comma separated values\n");
		fprintf(stderr, "       Options:\n");
		fprintf(stderr, "         -l: list sheets of Excel file, don't display content\n");
		fprintf(stderr, "         <sheet>: display only this sheet\n");
		fprintf(stderr, "       Output:\n");
		fprintf(stderr, "         %c is used to quote strings\n", stringSeparator);
		fprintf(stderr, "         LF is used to identify end of lines\n");
		fprintf(stderr, "         %s is used to identify end of field\n", fieldSeparator);
		fprintf(stderr, "\n");
		return EXIT_FAILURE;
	}

	// open workbook, choose standard conversion
	pWB = xls_open(argv[1], "iso-8859-15//TRANSLIT");
	if (!pWB) {
		fprintf(stderr, "File not found");
		fprintf(stderr, "\n");
		return EXIT_FAILURE;
	}

	// check if the requested sheet (if any) exists
	if ((argc >= 3) && (strcmp(argv[2], "-l") != 0)) {
		for (i = 0; i < pWB->sheets.count; i++) {
			if (strcmp(argv[2], pWB->sheets.sheet[i].name) == 0) {
				break;
			}
		}

		if (i == pWB->sheets.count) {
			fprintf(stderr, "Sheet not found");
			fprintf(stderr, "\n");
			return EXIT_FAILURE;
		}
	}

	// process all sheets
	for (i = 0; i < pWB->sheets.count; i++) {
		int isFirstLine = 1;

		// check if this the sheet we want
		if (argc >= 3) {
			if (strcmp(argv[2], "-l") == 0) {
				printf("%s\n", pWB->sheets.sheet[i].name);
				continue;
			}
			if (strcmp(argv[2], pWB->sheets.sheet[i].name) != 0) {
				continue;
			}
		}

		// open and parse the sheet
		pWS = xls_getWorkSheet(pWB, i);
		xls_parseWorkSheet(pWS);

		// process all rows of the sheet
		for (cellRow = 0; cellRow <= pWS->rows.lastrow; cellRow++) {
			int isFirstCol = 1;
			row = xls_row(pWS, cellRow);

			// process cells
			if (!isFirstLine) {
				printf(lineSeparator);
			} else {
				isFirstLine = 0;
			}

			for (cellCol = 0; cellCol <= pWS->rows.lastcol; cellCol++) {
				xlsCell *cell = xls_cell(pWS, cellRow, cellCol);

				if ((!cell) || (cell->ishiden)) {
					continue;
				}

				if (!isFirstCol) {
					printf(fieldSeparator);
				} else {
					isFirstCol = 0;
				}

				// display the colspan as only one cell, but reject rowspans (they can't be converted to CSV)
				if (cell->rowspan > 1) {
					fprintf(stderr, "%d,%d: rowspan=%i", cellCol, cellRow, cell->rowspan);
					fprintf(stderr, "\n");
					return EXIT_FAILURE;
				}

				// display the value of the cell (either numeric or string)
				if (cell->id == 0x27e || cell->id == 0x0BD || cell->id == 0x203) {
					OutputNumber(cell->d);
				} else if (cell->id == 0x06) // formula
				{
					if (cell->l == 0) // its a number
					{
						OutputNumber(cell->d);
					} else {
						if (cell->str == "bool") // its boolean, and test cell->d
						{
							OutputString((int) cell->d ? "true" : "false");
						} else if (cell->str == "error") // formula is in error
						{
							OutputString("*error*");
						} else // ... cell->str is valid as the result of a string formula.
						{
							OutputString(cell->str);
						}
					}
				} else if (cell->str != NULL) {
					OutputString(cell->str);
				} else {
					OutputString("");
				}
			}
		}
	}

	xls_close(pWB);
	return EXIT_SUCCESS;
}

// Output a CSV String (between double quotes)
// Escapes (doubles) " and \ characters
static void OutputString(const char *string) {
	const char *str;

	printf("%c", stringSeparator);
	for (str = string; *str; str++) {
		if (*str == stringSeparator) {
			printf("%c%c", stringSeparator, stringSeparator);
		} else if (*str == '\\') {
			printf("\\\\");
		} else {
			printf("%c", *str);
		}
	}
	printf("%c", stringSeparator);
}

// Output a CSV Number
static void OutputNumber(const double number) {
	printf("%.15g", number);
}
