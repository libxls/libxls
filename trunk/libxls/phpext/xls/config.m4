dnl $Id: config.m4,v 1.1.1.1 2003-09-02 07:21:43 kvmurom Exp $
dnl config.m4 for extension xls

dnl Comments in this file start with the string 'dnl'.
dnl Remove where necessary. This file will not work
dnl without editing.

dnl If your extension references something external, use with:

PHP_ARG_WITH(xls, for xls support,
 Make sure that the comment is aligned:
[ --with-xls             Include xls support])

dnl Otherwise use enable:

#PHP_ARG_ENABLE(xls, whether to enable xls support,
# Make sure that the comment is aligned:
# [--enable-xls           Enable xls support])

if test "$PHP_XLS" != "no"; then
  dnl Write more examples of tests here...

  dnl # --with-xls -> check with-path
  SEARCH_PATH="/usr/local/libxls /usr/local /usr"     # you might want to change this
  SEARCH_FOR="/include/libxls/xls.h"  # you most likely want to change this
   if test -r $PHP_XLS/$SEARCH_FOR; then # path given as parameter
     XLS_DIR=$PHP_XLS
   else # search default path list
     AC_MSG_CHECKING([for xls files in default path])
     for i in $SEARCH_PATH ; do
       if test -r $i/$SEARCH_FOR; then
         XLS_DIR=$i
         AC_MSG_RESULT(found in $i)
       fi
     done
   fi
  
   if test -z "$XLS_DIR"; then
    AC_MSG_RESULT([not found])
    AC_MSG_ERROR([Please reinstall the xls distribution])
   fi

  dnl # --with-xls -> add include path
   PHP_ADD_INCLUDE($XLS_DIR/include)

  dnl # --with-xls -> check for lib and symbol presence
   LIBNAME=xls # you may want to change this
   LIBSYMBOL=xls # you most likely want to change this 

   PHP_CHECK_LIBRARY($LIBNAME,$LIBSYMBOL,
   [
     PHP_ADD_LIBRARY_WITH_PATH($LIBNAME, $XLS_DIR/lib, XLS_SHARED_LIBADD)
     AC_DEFINE(HAVE_XLSLIB,1,[ ])
   ],[
     AC_MSG_ERROR([wrong xls lib version or lib not found])
   ],[
     -L$XLS_DIR/lib -lm -liconv
   ])

   PHP_ADD_LIBRARY_WITH_PATH(iconv, /usr/local/lib, XLS_SHARED_LIBADD)
  
   PHP_SUBST(XLS_SHARED_LIBADD)

  PHP_NEW_EXTENSION(xls, xls.c, $ext_shared)
fi
