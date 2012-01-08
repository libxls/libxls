AM_CPPFLAGS = -I$(top_srcdir)/include

noinst_PROGRAMS = test test2

test_SOURCES = test.c
test_LDADD = $(top_builddir)/src/libxlsreader.la

test2_SOURCES = test2.c
test2_LDADD = $(top_builddir)/src/libxlsreader.la

# Extra files to distribute
EXTRA_DIST = files/test2.xls

check_PROGRAMS = test
TESTS = test
