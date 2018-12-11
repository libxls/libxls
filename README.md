[![Build Status](https://travis-ci.org/evanmiller/libxls.svg?branch=master)](https://travis-ci.org/evanmiller/libxls)
[![Build status](https://ci.appveyor.com/api/projects/status/4ais4ilmbhuu605c?svg=true)](https://ci.appveyor.com/project/evanmiller/libxls)

libxls - Read XLS files from C
==

*Shhhhh don't tell. This repo will soon house the official libxls.*

This is a copy/fork/successor of libxls, a C library for reading Excel files in
the nasty old binary OLE format. Changes from this fork compared to the [original](https://sourceforge.net/projects/libxls/):

* Hosted on GitHub (hooray!)
* New in-memory parsing API
* Internals rewritten to return errors instead of exiting
* Heavily fuzz-tested with clang's libFuzzer, fixing many memory leaks and *cough* CVEs
* Improved compatibility with C++
* Continuous integration tests on Mac, Linux, and Windows
* Lots of other small fixes, see the commit history

The [C API](include/xls.h) is pretty simple, this will get you started:

```{C}
xls_error_t error = LIBXLS_OK;
xlsWorkBook *wb = xls_open_file("/path/to/finances.xls", "UTF-8", &error);
if (wb == NULL) {
    printf("Error reading file: %s\n", xls_getError(error));
} else {
    for (int i=0; i<wb->sheets.count; i++) { // sheets
        xl_WorkSheet *work_sheet = xls_getWorkSheet(work_book, i);
        error = xls_parseWorkSheet(work_sheet);
        for (int j=0; j<=work_sheet->rows.lastrow; j++) { // rows
            xlsRow *row = xls_row(work_sheet, j);
            for (int k=0; k<=work_sheet->rows.lastcol; k++) { // columns
                xlsCell *cell = &row->cells.cell[k];
                // do something with cell
                if (cell->id == XLS_RECORD_FORMULA) { // formula
                } else if (cell->l == 0) { // its a number
                    ... use cell->d
                } else {
                    if(cell->str == "bool") // its boolean, and test cell->d > 0.0 for true
                    if(cell->str == "error") // formula is in error
                    else ... cell->str is valid as the result of a string formula.
                }
            }
        }
        xls_close_WS(work_sheet);
    }
    xls_close_WB(wb);
}
```

The library also includes a CLI tool for converting Excel files to CSV:

    ./xls2csv /path/to/file.xls

Libxls should run fine on both little-endian and big-endian systems, but if not
please open an issue.

If you want to hack on the source, you should first familiarize yourself with the [Microsoft Excel File Format](http://sc.openoffice.org/excelfileformat.pdf) as well as [Coumpound Document file format](http://sc.openoffice.org/compdocfileformat.pdf) (documentation provided by the nice folks at OpenOffice.org).

Installation
---

If you want a stable version, head back to [Sourceforge](https://sourceforge.net/projects/libxls/files/) and download 1.4.0. Otherwise see [INSTALL](INSTALL), or here's the tl;dr:

```
./configure
make
make install
```

Once the dust settles on this repo, I'll mark a 1.5 release. But don't tell anyone.
