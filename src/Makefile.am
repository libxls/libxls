INCLUDES = -I$(top_srcdir)/include -I$(top_builddir)/include

libxls_sources =	\
	xlstool.c	\
	ole.c		\
	xls.c		

lib_LTLIBRARIES = libxls.la

LDADD = $(LIBXLS_LIBS)

libxls_la_LDFLAGS = -version-info $(VERSION_INFO)

libxls_la_SOURCES = $(libxls_sources) version.c