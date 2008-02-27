INCLUDES = -I$(top_srcdir)/src -I$(top_srcdir) -I../include

noinst_PROGRAMS = test test2 xls2csv
test_LDADD = $(top_builddir)/src/libxlsreader.la $(LIBXLS_LIBS) -L$(top_builddir)/src -lxlsreader -liconv
test2_LDADD = $(top_builddir)/src/libxlsreader.la $(LIBXLS_LIBS) -L$(top_builddir)/src -lxlsreader -liconv
xls2csv_LDADD = $(top_builddir)/src/libxlsreader.la $(LIBXLS_LIBS) -L$(top_builddir)/src -lxlsreader -liconv

test_LDFLAGS = 
AM_CFLAGS = 
AM_CXXFLAGS = 
AM_FFLAGS = 
SUBDIRS = 
