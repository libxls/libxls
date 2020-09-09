AC_INIT([libxls],[1.6.1],[emmiller@gmail.com], [libxls], [https://github.com/libxls/libxls])
AC_CONFIG_SRCDIR([test/test.c])

LIBXLS_MAJOR_VERSION=1
LIBXLS_MINOR_VERSION=6
LIBXLS_MICRO_VERSION=1

VERSION=$LIBXLS_MAJOR_VERSION.$LIBXLS_MINOR_VERSION.$LIBXLS_MICRO_VERSION
# if libxls ever goes to 2.0 you will need to change the following line to
# start at whatever the ending value was in the 1.x series E.g. if the last
# minor release was 1.9 then the value of "current" was 11 so the line should
# be changed to `expr 12 + $LIBXLS_MINOR_VERSION`
VERSION_INFO=`expr 2 + $LIBXLS_MINOR_VERSION`:$LIBXLS_MICRO_VERSION

AC_SUBST(VERSION_INFO)
AC_SUBST(LIBXLS_MAJOR_VERSION)
AC_SUBST(LIBXLS_MINOR_VERSION)
AC_SUBST(LIBXLS_MICRO_VERSION)

# Define those values into C code
AC_DEFINE_UNQUOTED([LIBXLS_MAJOR_VERSION],$LIBXLS_MAJOR_VERSION,[Major version])
AC_DEFINE_UNQUOTED([LIBXLS_MINOR_VERSION],$LIBXLS_MINOR_VERSION,[Minor version])
AC_DEFINE_UNQUOTED([LIBXLS_MICRO_VERSION],$LIBXLS_MICRO_VERSION,[Micro version])

AM_INIT_AUTOMAKE([foreign subdir-objects])
AM_SILENT_RULES([yes])

LT_INIT([win32-dll disable-static])

AC_CANONICAL_BUILD
AC_CANONICAL_HOST

AC_CONFIG_HEADERS([config.h])
AC_PREFIX_DEFAULT(/usr/local)

AC_PROG_INSTALL
AC_PROG_CC
AC_PROG_CC_C99
AC_PROG_CXX
AX_CXX_COMPILE_STDCXX_11([], [optional])

AS_IF([test "x$HAVE_CXX11" != x1], [
    AC_MSG_NOTICE([---])
    AC_MSG_NOTICE([No C++11 support, will skip C++ tests.])
    AC_MSG_NOTICE([---])
])
AM_CONDITIONAL([HAVE_CXX11], [test x$HAVE_CXX11 = x1])

AC_CHECK_FUNCS([strdup wcstombs_l])
AC_CHECK_HEADERS([wchar.h xlocale.h])
AC_FUNC_MALLOC
AC_FUNC_REALLOC
AC_TYPE_SIZE_T

AC_ARG_ENABLE([fuzz-testing], AS_HELP_STRING([--enable-fuzz-testing], ["Enable fuzz testing (requires Clang 6 or later)"]), [
   AC_MSG_CHECKING([whether $CC accepts -fsanitize=fuzzer])
   tmp_saved_flags=$[]_AC_LANG_PREFIX[]FLAGS
   _AC_LANG_PREFIX[]FLAGS="$[]_AC_LANG_PREFIX[]FLAGS -fsanitize=fuzzer"
   AC_LINK_IFELSE([AC_LANG_PROGRAM()],
       [
	    AC_MSG_RESULT(yes)
        fuzzer=yes],
	    AC_MSG_RESULT(no)
       AC_MSG_FAILURE([-fsanitize=fuzzer not supported (Required with --enable-fuzz-testing)]))
   _AC_LANG_PREFIX[]FLAGS=$tmp_saved_flags
], [fuzzer=no])
AM_CONDITIONAL([FUZZER_ENABLED], test "x$fuzzer" = "xyes")

AC_ARG_VAR([LIB_FUZZING_ENGINE], [Location of prebuilt fuzzing engine library])
AC_SUBST([LIB_FUZZING_ENGINE])

AS_CASE([$host],
        [*mingw*|*cygwin*], [EXTRA_LDFLAGS="-no-undefined"], [EXTRA_LDFLAGS=""])
AC_SUBST([EXTRA_LDFLAGS])

AM_ICONV

m4_ifdef([PKG_INSTALLDIR], [
    PKG_INSTALLDIR
], [
    AC_ARG_WITH([pkgconfigdir],
        [AS_HELP_STRING([--with-pkgconfigdir],
                        [pkg-config installation directory @<:@'${libdir}/pkgconfig'@:>@])],,
        [with_pkgconfigdir='${libdir}/pkgconfig'])
    AC_SUBST([pkgconfigdir], [$with_pkgconfigdir])
])

LIBXLS_LIBS="$LTLIBICONV $LIBS"

AC_SUBST(LIBXLS_LIBS)

AC_CONFIG_FILES([Makefile libxls.pc])
AC_OUTPUT

AC_MSG_NOTICE([---])
AC_MSG_NOTICE([catdoc users: avoid xls2csv conflicts, use ./configure --program-prefix=lib to get libxls2csv])
AC_MSG_NOTICE([---])

AC_MSG_RESULT([
Configuration:

C compiler: $CC
C++ compiler: $CXX
CFLAGS: $CFLAGS
CXXFLAGS: $CXXFLAGS
LIBS: $LIBXLS_LIBS

Host: $host
Extra ld flags: $EXTRA_LDFLAGS])
