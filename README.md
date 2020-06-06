[![Build Status](https://travis-ci.org/libxls/libxls.svg?branch=master)](https://travis-ci.org/libxls/libxls)
[![Build Status](https://ci.appveyor.com/api/projects/status/3nx26kfmy2y0efsi?svg=true)](https://ci.appveyor.com/project/evanmiller/libxls-252ki)
[![Fuzzing Status](https://oss-fuzz-build-logs.storage.googleapis.com/badges/libxls.svg)](https://bugs.chromium.org/p/oss-fuzz/issues/list?sort=-opened&can=1&q=proj:libxls)

libxls - Read XLS files from C
==

This is libxls, a C library for reading Excel files in the nasty old binary OLE
format, plus a command-line tool for converting XLS to CSV (named, appropriately
enough, `xls2csv`).

After several years of neglect, libxls is under new management as of the 1.5.x
series. Head over to [releases](https://github.com/libxls/libxls/releases) to
get the latest stable version of libxls 1.5, which fixes *many* security
vulnerabilities found in libxls 1.4 and earlier.

Libxls 1.5 also includes new APIs for parsing files stored in memory buffers,
and returns errors instead of exiting upon encountering malformed input. If you
find a bug, please file it on the [GitHub issue tracker](https://github.com/libxls/libxls/issues).

Changes to libxls since 1.4:

* Hosted on GitHub (hooray!)
* New in-memory parsing API (see `xls_open_buffer`)
* Internals rewritten to return errors instead of exiting
* Heavily fuzz-tested with clang's libFuzzer, fixing many memory leaks and CVEs
* Improved compatibility with C++
* Continuous integration tests on Mac, Linux, and Windows
* Lots of other small fixes, see the commit history

The [C API](include/xls.h) is pretty simple, this will get you started:

```c
xls_error_t error = LIBXLS_OK;
xlsWorkBook *wb = xls_open_file("/path/to/finances.xls", "UTF-8", &error);
if (wb == NULL) {
    printf("Error reading file: %s\n", xls_getError(error));
    exit(1);
}
for (int i=0; i<wb->sheets.count; i++) { // sheets
    xlsWorkSheet *work_sheet = xls_getWorkSheet(work_book, i);
    error = xls_parseWorkSheet(work_sheet);
    for (int j=0; j<=work_sheet->rows.lastrow; j++) { // rows
        xlsRow *row = xls_row(work_sheet, j);
        for (int k=0; k<=work_sheet->rows.lastcol; k++) { // columns
            xlsCell *cell = &row->cells.cell[k];
            // do something with cell
            if (cell->id == XLS_RECORD_BLANK) {
                // do something with a blank cell
            } else if (cell->id == XLS_RECORD_NUMBER) {
               // use cell->d, a double-precision number
            } else if (cell->id == XLS_RECORD_FORMULA) {
                if (strcmp(cell->str, "bool") == 0) {
                    // its boolean, and test cell->d > 0.0 for true
                } else if (strcmp(cell->str, "error") == 0) {
                    // formula is in error
                } else {
                    // cell->str is valid as the result of a string formula.
                }
            } else if (cell->str != NULL) {
                // cell->str contains a string value
            }
        }
    }
    xls_close_WS(work_sheet);
}
xls_close_WB(wb);
```

The library also includes a CLI tool for converting Excel files to CSV:

    ./xls2csv /path/to/file.xls

The man page for `xls2csv` has more details.

Libxls should run fine on both little-endian and big-endian systems, but if not
please open an [issue](https://github.com/libxls/libxls/issues/new).

If you want to hack on the source, you should first familiarize yourself with
the [Microsoft Excel File Format](http://sc.openoffice.org/excelfileformat.pdf)
as well as [Compound Document file
format](http://sc.openoffice.org/compdocfileformat.pdf) (documentation provided
by the nice folks at OpenOffice.org).

Installation
---

If you want a stable version, check out the
[Releases](https://github.com/libxls/libxls/releases) section, which has copies of everything
you'll find in [Sourceforge](https://sourceforge.net/projects/libxls/files/),
and download version 1.5.0 or later.

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

That will generate all the supporting files. It assumes autotools is already
installed on the system (and also expects Autoconf Archive to be present).

Language bindings
---

If C is not your cup of tea, you can make use of libxls in several other languages, including:

* [Haskell](https://hackage.haskell.org/package/xls)
* [R](https://readxl.tidyverse.org)
* [Rust](https://github.com/evanmiller/rust-xls)
