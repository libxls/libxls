[![Build Status](https://travis-ci.org/libxls/libxls.svg?branch=master)](https://travis-ci.org/libxls/libxls)
[![Build status](https://ci.appveyor.com/api/projects/status/4ais4ilmbhuu605c?svg=true)](https://ci.appveyor.com/project/evanmiller/libxls)

libxls - Read XLS files from C
==

This is libxls, a C library for reading Excel files in the nasty old binary OLE
format. **We are in the process of preparing a 1.5 release and moving the project
over from [SourceForge](https://sourceforge.net/projects/libxls/).** If you need
a stable release, head back to SourceForge, or see the
[releases](https://github.com/libxls/libxls/releases) section, which currently
has copies of everything from SourceForge.

Please note that the current stable releases have several public security
vulnerabilities. We'll have them fixed in 1.5. Hang tight.

Changes since 1.4:

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
please open an [issue](https://github.com/libxls/libxls/issues/new).

If you want to hack on the source, you should first familiarize yourself with the [Microsoft Excel File Format](http://sc.openoffice.org/excelfileformat.pdf) as well as [Coumpound Document file format](http://sc.openoffice.org/compdocfileformat.pdf) (documentation provided by the nice folks at OpenOffice.org).

Installation
---

If you want a stable version, check out the
[Releases](https://github.com/libxls/libxls/releases) section, which has copies of everything
you'll find in [Sourceforge](https://sourceforge.net/projects/libxls/files/),
and download version 1.4.0.

For full instructions see [INSTALL](INSTALL), or here's the tl;dr:

To install a stable release:

```
./configure
make
make install
```

If you've cloned the git repository, you'll need to run this first:

```
./bootstrap
```

That will generate all the supporting files (assuming autotools is already
present on the system).
