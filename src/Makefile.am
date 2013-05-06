AM_CPPFLAGS = -I$(top_srcdir)/include

bin_PROGRAMS = xls2csv

lib_LTLIBRARIES = libxlsreader.la

xls2csv_SOURCES =	\
	xls2csv.c		\
	getopt.c

xls2csv_LDADD = libxlsreader.la

libxlsreader_la_SOURCES =	\
	xlstool.c				\
	endian.c				\
	ole.c					\
	xls.c

libxlsreader_la_LDFLAGS = -version-info $(VERSION_INFO) $(LIBXLS_LIBS)

