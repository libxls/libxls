AM_CPPFLAGS = -I$(top_srcdir)/include -I$(top_builddir)/include

bin_PROGRAMS = xls2csv

lib_LTLIBRARIES = libxls.la

xls2csv_sources =       \
	xls2csv.c

libxls_sources =	\
	xlstool.c	\
	ole.c		\
	xls.c		

LDADD = $(LIBXLS_LIBS)

libxls_la_LDFLAGS = -version-info $(VERSION_INFO)

libxls_la_SOURCES = $(libxls_sources) version.c
