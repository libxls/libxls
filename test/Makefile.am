AM_CPPFLAGS = -I$(top_srcdir)/include -I$(top_srcdir) -I../include

noinst_PROGRAMS = test

test_SOURCES = \
	test.c

test_LDADD = $(top_builddir)/src/libxls.la $(LIBXLS_LIBS) -L$(top_builddir)/src -lxls

# Extra files to distribute
EXTRA_DIST = files/test2.xls
