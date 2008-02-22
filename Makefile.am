## Process this file with automake to create Makefile.in.
AUTOMAKE_OPTIONS = foreign dist-zip
SUBDIRS = include src test 

# Extra files to distribute
EXTRA_DIST = doc

dist-hook:
	rm -rf `find $(distdir)/doc -name CVS`