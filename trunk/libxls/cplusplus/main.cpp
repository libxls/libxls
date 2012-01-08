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
 * Copyright 2012 David Hoerl
 *
 */

#include <iostream>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include <assert.h>

#include "XlsReader.h"

using namespace xls;
using namespace std;

int main(int argc, char *argv[])
{
#warning Provide a hard coded file path
	string s = string("/tmp/xls.xls");
	WorkBook foo(s);
	
	cellContent cell = foo.GetCell(0, 1, 2);
	foo.ShowCell(cell);
	
	foo.InitIterator(0);
	while(true) {
		cellContent c = foo.GetNextCell();
		if(c.type == cellBlank) break;
		foo.ShowCell(c);
	}
}