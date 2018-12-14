## Process this file with automake to create Makefile.in.
AUTOMAKE_OPTIONS = foreign dist-zip

includedir = $(prefix)

AM_CPPFLAGS = -Iinclude

nobase_include_HEADERS = \
	include/xls.h \
	include/libxls/brdb.c.h	\
	include/libxls/brdb.h \
	include/libxls/endian.h	\
	include/libxls/ole.h \
	include/libxls/xlsstruct.h \
	include/libxls/xlstypes.h \
	include/libxls/xlstool.h

bin_PROGRAMS = xls2csv
noinst_PROGRAMS = test_libxls test2_libxls

lib_LTLIBRARIES = libxlsreader.la

xls2csv_SOURCES = \
	src/xls2csv.c \
	src/getopt.c

xls2csv_LDADD = libxlsreader.la

libxlsreader_la_SOURCES = \
	src/xlstool.c \
	src/endian.c \
	src/ole.c \
	src/xls.c

libxlsreader_la_LDFLAGS = -version-info $(VERSION_INFO) $(LIBXLS_LIBS) @EXTRA_LDFLAGS@
libxlsreader_la_CFLAGS = -Wall -Wextra -Wstrict-prototypes -Wno-unused-parameter -Werror -pedantic-errors

if FUZZER_ENABLED
libxlsreader_la_CFLAGS += -fsanitize=fuzzer-no-link
libxlsreader_la_LDFLAGS += -fsanitize=fuzzer
endif

test_libxls_SOURCES = test/test.c
test_libxls_LDADD = libxlsreader.la

test2_libxls_SOURCES = test/test2.c
test2_libxls_LDADD = libxlsreader.la

check_PROGRAMS = test_libxls test2_libxls
TESTS = test_libxls

if FUZZER_ENABLED
noinst_PROGRAMS += fuzz_xls

fuzz_xls_SOURCES = fuzz/fuzz_xls.c
fuzz_xls_CFLAGS = -fsanitize=fuzzer
fuzz_xls_LDADD = libxlsreader.la -lstdc++
endif

# Extra files to distribute
EXTRA_DIST = doc test/files/test2.xls

dist-hook:
	rm -rf `find $(distdir)/doc -name CVS`

