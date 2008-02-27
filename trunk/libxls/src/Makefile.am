INCLUDES = -I$(top_srcdir)/include -I$(top_builddir)/include

libxls_sources =	\
	xlstool.c	\
	ole.c		\
	xls.c		

lib_LTLIBRARIES = libxlsreader.la

LDADD = $(LIBXLS_LIBS)

libxlsreader_la_LDFLAGS = -version-info $(VERSION_INFO)

libxlsreader_la_SOURCES = $(libxls_sources) version.c
