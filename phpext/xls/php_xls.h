/*
  +----------------------------------------------------------------------+
  | PHP Version 4                                                        |
  +----------------------------------------------------------------------+
  | Copyright (c) 1997-2003 The PHP Group                                |
  +----------------------------------------------------------------------+
  | This source file is subject to version 3.0 of the PHP license,       |
  | that is bundled with this package in the file LICENSE, and is        |
  | available through the world-wide-web at the following url:           |
  | http://www.php.net/license/3_0.txt.                                  |
  | If you did not receive a copy of the PHP license and are unable to   |
  | obtain it through the world-wide-web, please send a note to          |
  | license@php.net so we can mail you a copy immediately.               |
  +----------------------------------------------------------------------+
  | Author:                                                              |
  +----------------------------------------------------------------------+
*/

/* $Id: php_xls.h,v 1.2 2003-11-14 10:59:45 tony2001 Exp $ */

#ifndef PHP_XLS_H
#define PHP_XLS_H

extern zend_module_entry xls_module_entry;
#define phpext_xls_ptr &xls_module_entry

#ifdef PHP_WIN32
#define PHP_XLS_API __declspec(dllexport)
#else
#define PHP_XLS_API
#endif

#ifdef ZTS
#include "TSRM.h"
#endif

#include <libxls/xls.h>

PHP_MINIT_FUNCTION(xls);
PHP_MSHUTDOWN_FUNCTION(xls);
PHP_RINIT_FUNCTION(xls);
PHP_RSHUTDOWN_FUNCTION(xls);
PHP_MINFO_FUNCTION(xls);

PHP_FUNCTION(xls_open);	
PHP_FUNCTION(xls_getcharset);
PHP_FUNCTION(xls_getsheetscount);
PHP_FUNCTION(xls_getsheetname);
PHP_FUNCTION(xls_getworksheet);
PHP_FUNCTION(xls_parseworksheet);
PHP_FUNCTION(xls_getcss);
PHP_FUNCTION(xls_fetch_worksheet);

#ifdef ZTS
#define XLS_G(v) TSRMG(xls_globals_id, zend_xls_globals *, v)
#else
#define XLS_G(v) (xls_globals.v)
#endif

#endif	/* PHP_XLS_H */


/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * End:
 * vim600: noet sw=4 ts=4 fdm=marker
 * vim<600: noet sw=4 ts=4
 */
