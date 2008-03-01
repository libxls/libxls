AM_CPPFLAGS = -I$(top_srcdir)/include -I$(top_builddir)/include

bin_PROGRAMS = xls2csv

lib_LTLIBRARIES = libxlsreader.la

xls2csv_SOURCES =       \
	xls2csv.c

xls2csv_LDADD = libxlsreader.la

libxlsreader_la_SOURCES =	\
	xlstool.c	\
	ole.c		\
	xls.c		\
	version.c

libxlsreader_la_LDFLAGS = -version-info $(VERSION_INFO) $(LIBXLS_LIBS)
