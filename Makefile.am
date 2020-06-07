## Process this file with automake to create Makefile.in.
AUTOMAKE_OPTIONS = foreign dist-zip

includedir = $(prefix)

AM_CPPFLAGS = -Iinclude

nobase_include_HEADERS = \
	include/xls.h \
	include/libxls/brdb.c.h	\
	include/libxls/brdb.h \
	include/libxls/endian.h	\
	include/libxls/locale.h	\
	include/libxls/ole.h \
	include/libxls/xlsstruct.h \
	include/libxls/xlstypes.h \
	include/libxls/xlstool.h

bin_PROGRAMS = xls2csv
noinst_PROGRAMS = test_libxls test2_libxls
noinst_HEADERS = cplusplus/XlsReader.h
check_PROGRAMS = test_libxls test2_libxls

if HAVE_CXX11
noinst_PROGRAMS += test_cpp
check_PROGRAMS += test_cpp
endif

dist_man1_MANS = man/xls2csv.man

pkgconfig_DATA = libxls.pc

lib_LTLIBRARIES = libxlsreader.la

xls2csv_SOURCES = \
	src/xls2csv.c

xls2csv_LDADD = libxlsreader.la

libxlsreader_la_SOURCES = \
	src/xlstool.c \
	src/endian.c \
	src/locale.c \
	src/ole.c \
	src/xls.c

libxlsreader_la_LDFLAGS = -version-info $(VERSION_INFO) $(LIBXLS_LIBS) @EXTRA_LDFLAGS@
libxlsreader_la_CFLAGS = -Wall -Wextra -Wstrict-prototypes -Wno-unused-parameter -pedantic-errors

if FUZZER_ENABLED
libxlsreader_la_CFLAGS += -fsanitize=fuzzer-no-link -fsanitize=address
libxlsreader_la_LDFLAGS += -fsanitize=fuzzer -fsanitize=address
endif

test_libxls_SOURCES = test/test.c
test_libxls_LDADD = libxlsreader.la

test2_libxls_SOURCES = test/test2.c
test2_libxls_LDADD = libxlsreader.la

if HAVE_CXX11
test_cpp_SOURCES = cplusplus/main.cpp cplusplus/XlsReader.cpp
test_cpp_LDADD = libxlsreader.la
endif

TESTS = test_libxls

EXTRA_DIST = README.md LICENSE test/files/test2.xls

distclean-local:
	-rm -f $(builddir)/test.htm

EXTRA_PROGRAMS = fuzz_xls

# Force C++ linking for fuzz target
nodist_EXTRA_fuzz_xls_SOURCES = dummy.cxx

fuzz_xls_SOURCES = fuzz/fuzz_xls.c
fuzz_xls_CFLAGS = -Wall -pedantic-errors -std=c99
fuzz_xls_LDADD = libxlsreader.la @LIB_FUZZING_ENGINE@
fuzz_xls_LDFLAGS = -static
