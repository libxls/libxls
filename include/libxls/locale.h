#ifdef HAVE_XLOCALE_H
#include <xlocale.h>
#else
#include <locale.h>
#endif

#if defined(_WIN32) || defined(WIN32) || defined(_WIN64) || defined(WIN64) || defined(WINDOWS)
typedef _locale_t xls_locale_t;
#else
typedef locale_t xls_locale_t;
#endif

xls_locale_t xls_createlocale(void);
void xls_freelocale(xls_locale_t locale);
size_t xls_wcstombs_l(char *restrict s, const wchar_t *restrict pwcs, size_t n, xls_locale_t loc);
