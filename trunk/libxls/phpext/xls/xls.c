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

/* $Id: xls.c,v 1.1.1.1 2003-09-02 07:21:43 kvmurom Exp $ */

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include "php.h"
#include "php_ini.h"
#include "ext/standard/info.h"
#include "php_xls.h"


#define s_xls_wb "WorkBook"
#define s_xls_ws "WorkSheet"

typedef struct {
	int link;
	int type;
} xls_res;

/* If you declare any globals in php_xls.h uncomment this:
ZEND_DECLARE_MODULE_GLOBALS(xls)
*/

/* True global resources - no need for thread safety here */
static int le_xls_wb;
static int le_xls_ws;

zend_class_entry *xls_class_entry;

xlsWorkBook* xls_find_workbook(pval *id TSRMLS_DC)
{
	int type;
	xlsWorkBook *wb;


	convert_to_long(id);
	wb = zend_list_find(Z_LVAL_P(id), &type); 

	if(!wb) { 
		php_error_docref(NULL TSRMLS_CC, E_WARNING, "Unable to find identifier %d", id);
		return NULL; 
	} 
	if (type!=le_xls_wb) { 
		php_error_docref(NULL TSRMLS_CC, E_WARNING, "Identifier is not WorkSheet");
		return NULL; 
	} 

	return wb;
}


xlsWorkSheet* xls_find_worksheet(pval *id TSRMLS_DC)
{
	int type;
	xlsWorkSheet *ws;

	convert_to_long(id);
	ws = zend_list_find(Z_LVAL_P(id), &type); 


	if(!ws) { 
		php_error_docref(NULL TSRMLS_CC, E_WARNING, "Unable to find identifier %d", id);
		return NULL; 
	} 
	if (type!=le_xls_ws) { 
		php_error_docref(NULL TSRMLS_CC, E_WARNING, "Identifier is not WorkSheet");
		return NULL; 
	} 
	return ws;
}


/* {{{ xls_functions[]
 *
 * Every user visible function must have an entry in xls_functions[].
 */
function_entry xls_functions[] = {
	PHP_FE(xls_open,NULL)		/* Open XLS file*/
	PHP_FE(xls_getcharset,NULL)
	PHP_FE(xls_getsheetscount,NULL)
	PHP_FE(xls_getsheetname,NULL)
	PHP_FE(xls_getworksheet,NULL)
	PHP_FE(xls_parseworksheet,NULL)
	PHP_FE(xls_getcss,NULL)
	PHP_FE(xls_fetch_worksheet,NULL)
	{NULL, NULL, NULL}	/* Must be the last line in xls_functions[] */
};
/* }}} */

/* {{{ xls_module_entry
 */
zend_module_entry xls_module_entry = {
#if ZEND_MODULE_API_NO >= 20010901
	STANDARD_MODULE_HEADER,
#endif
	"xls",
	xls_functions,
	PHP_MINIT(xls),
	PHP_MSHUTDOWN(xls),
	PHP_RINIT(xls),		/* Replace with NULL if there's nothing to do at request start */
	PHP_RSHUTDOWN(xls),	/* Replace with NULL if there's nothing to do at request end */
	PHP_MINFO(xls),
#if ZEND_MODULE_API_NO >= 20010901
	"0.1", /* Replace with version number for your extension */
#endif
	STANDARD_MODULE_PROPERTIES
};
/* }}} */

#ifdef COMPILE_DL_XLS
ZEND_GET_MODULE(xls)
#endif

/* {{{ PHP_INI
 */
/* Remove comments and fill if you need to have entries in php.ini
PHP_INI_BEGIN()
    STD_PHP_INI_ENTRY("xls.global_value",      "42", PHP_INI_ALL, OnUpdateLong, global_value, zend_xls_globals, xls_globals)
    STD_PHP_INI_ENTRY("xls.global_string", "foobar", PHP_INI_ALL, OnUpdateString, global_string, zend_xls_globals, xls_globals)
PHP_INI_END()
*/
/* }}} */

/* {{{ php_xls_init_globals
 */
/* Uncomment this function if you have INI entries
static void php_xls_init_globals(zend_xls_globals *xls_globals)
{
	xls_globals->global_value = 0;
	xls_globals->global_string = NULL;
}
*/
/* }}} */

void xls_destructor(zend_rsrc_list_entry *rsrc TSRMLS_DC)
{
//	efree(rsrc->ptr);
//	zend_printf("Destructed resurce #%i\n",rsrc->type);
}


/* {{{ PHP_MINIT_FUNCTION
 */
PHP_MINIT_FUNCTION(xls)
{
	/* If you have INI entries, uncomment these lines 
	ZEND_INIT_MODULE_GLOBALS(xls, php_xls_init_globals, NULL);
	REGISTER_INI_ENTRIES();
	*/
	le_xls_wb = zend_register_list_destructors_ex(xls_destructor, NULL, s_xls_wb,module_number);
	le_xls_ws = zend_register_list_destructors_ex(xls_destructor, NULL, s_xls_ws,module_number);
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_MSHUTDOWN_FUNCTION
 */
PHP_MSHUTDOWN_FUNCTION(xls)
{
	/* uncomment this line if you have INI entries
	UNREGISTER_INI_ENTRIES();
	*/
	return SUCCESS;
}
/* }}} */

/* Remove if there's nothing to do at request start */
/* {{{ PHP_RINIT_FUNCTION
 */
PHP_RINIT_FUNCTION(xls)
{
	return SUCCESS;
}
/* }}} */

/* Remove if there's nothing to do at request end */
/* {{{ PHP_RSHUTDOWN_FUNCTION
 */
PHP_RSHUTDOWN_FUNCTION(xls)
{
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_MINFO_FUNCTION
 */
PHP_MINFO_FUNCTION(xls)
{
	php_info_print_table_start();
	php_info_print_table_row(2, "xls support", "enabled");
	php_info_print_table_row(2, "xls version",libxls_version);
	php_info_print_table_end();

	/* Remove comments if you have entries in php.ini
	DISPLAY_INI_ENTRIES();
	*/
}
/* }}} */


/* Remove the following function when you have succesfully modified config.m4
   so that your module can be compiled into PHP, it exists only for testing
   purposes. */

/* Every user-visible function in PHP should document itself in the source */
/* {{{ proto string confirm_xls_compiled(string arg)
   Return a string to confirm that the module is compiled in */
PHP_FUNCTION(xls_open)
{
	char *file = NULL;
	char *charset = NULL;
	int file_len, charset_len;
	int id;
	xlsWorkBook* pWB;

	if (zend_parse_parameters(ZEND_NUM_ARGS() TSRMLS_CC, "ss", &file, &file_len, &charset,&charset_len ) == FAILURE) {
		return;
	}
	pWB=xls_open(file,charset);
	if (pWB) {
		id = zend_list_insert(pWB,le_xls_wb);
		RETURN_RESOURCE(id);
	} else {
		php_printf("%s(): Could not open file %s", get_active_function_name(TSRMLS_C), file);
		RETURN_FALSE;
	}

}

PHP_FUNCTION(xls_getcharset)
{
	pval *id;
	xlsWorkBook *wb;
	char* charset;

	if (ZEND_NUM_ARGS() != 1 || zend_get_parameters(ht, 1, &id)==FAILURE) {
		WRONG_PARAM_COUNT;
	}

	wb=xls_find_workbook(id TSRMLS_CC);
	if(wb==NULL) RETURN_FALSE;

	charset=wb->charset;
	RETURN_STRING(charset,1);
}

PHP_FUNCTION(xls_getsheetscount)
{
	pval *id;
	xlsWorkBook *wb;

	if (ZEND_NUM_ARGS() != 1 || zend_get_parameters(ht, 1, &id)==FAILURE) {
		WRONG_PARAM_COUNT;
	}

	wb=xls_find_workbook(id TSRMLS_CC);
	if(wb==NULL) RETURN_FALSE;

	RETURN_LONG(wb->sheets.count);
}

PHP_FUNCTION(xls_getsheetname)
{
	pval *id,*sheet;
	xlsWorkBook *wb;

	if (ZEND_NUM_ARGS() != 2 || zend_get_parameters(ht, 2, &id,&sheet)==FAILURE) {
		WRONG_PARAM_COUNT;
	}

	wb=xls_find_workbook(id TSRMLS_CC);
	if(wb==NULL) RETURN_FALSE;

    	RETVAL_STRING(wb->sheets.sheet[Z_LVAL_P(sheet)].name,1);
}


PHP_FUNCTION(xls_getworksheet)
{
	pval *id,*sheet;
	xlsWorkBook *wb;
	xlsWorkSheet *ws;

	if (ZEND_NUM_ARGS() != 2 || zend_get_parameters(ht, 2, &id,&sheet)==FAILURE) {
		WRONG_PARAM_COUNT;
	}

	wb=xls_find_workbook(id TSRMLS_CC);

	ws=xls_getWorkSheet(wb,Z_LVAL_P(sheet));

	if(ws==NULL) RETURN_FALSE;

	RETURN_RESOURCE(zend_list_insert(ws,le_xls_ws));
}

PHP_FUNCTION(xls_parseworksheet)
{
	pval *id;
	xlsWorkSheet *ws;

	if (ZEND_NUM_ARGS() != 1 || zend_get_parameters(ht, 1, &id)==FAILURE) {
		WRONG_PARAM_COUNT;
	}

	ws=xls_find_worksheet(id TSRMLS_CC);
	if (ws==NULL) RETURN_FALSE;

	xls_parseWorkSheet(ws);
	RETURN_TRUE
}

PHP_FUNCTION(xls_getcss)
{
	pval *id;
	xlsWorkBook *wb;
	char*	css;

	if (ZEND_NUM_ARGS() != 1 || zend_get_parameters(ht, 1, &id)==FAILURE) {
		WRONG_PARAM_COUNT;
	}

	wb=xls_find_workbook(id TSRMLS_CC);
	if(wb==NULL) RETURN_FALSE;
                                    
	css=xls_getCSS(wb);
	RETURN_STRING(css,1);
}

PHP_FUNCTION(xls_fetch_worksheet)
{
	zval *row;
	zval *cell;
	zval *cells;
	zval *rows;
	zval *arr,*arr2;

	pval *id;
	xlsWorkSheet *ws;
	int i,t;


	if (ZEND_NUM_ARGS() != 1 || zend_get_parameters(ht, 1, &id)==FAILURE) {
		WRONG_PARAM_COUNT;
	}

	ws=xls_find_worksheet(id TSRMLS_CC);

	if (ws==NULL) RETURN_FALSE;

	MAKE_STD_ZVAL(row);	
	object_init(row); 
	add_property_long(row,"lastcol",ws->rows.lastcol);
	add_property_long(row,"lastrow",ws->rows.lastrow);


	MAKE_STD_ZVAL(arr);
	array_init(arr);
	for (i=0;i<=ws->rows.lastrow;i++)	
	{
		MAKE_STD_ZVAL(rows);
		object_init(rows); 
		add_property_long(rows,"index",ws->rows.row[i].index);
		add_property_long(rows,"fcell",ws->rows.row[i].fcell);
		add_property_long(rows,"lcell",ws->rows.row[i].lcell);
		add_property_long(rows,"height",ws->rows.row[i].height);
		add_property_long(rows,"flags",ws->rows.row[i].flags);
		add_property_long(rows,"xf",ws->rows.row[i].xf);
		add_property_long(rows,"xfflags",ws->rows.row[i].xfflags);

//---------------------------------------------------------------------------------

	MAKE_STD_ZVAL(cell);	
	object_init(cell); 
	add_property_long(cell,"count",ws->rows.lastcol);

	MAKE_STD_ZVAL(arr2);
	array_init(arr2);
	for (t=0;t<=ws->rows.lastcol;t++)	
	{
		MAKE_STD_ZVAL(cells);
		object_init(cells); 
		add_property_long(cells,"id",ws->rows.row[i].cells.cell[t].id);
		add_property_long(cells,"row",ws->rows.row[i].cells.cell[t].row);
		add_property_long(cells,"col",ws->rows.row[i].cells.cell[t].col);
		add_property_long(cells,"xf",ws->rows.row[i].cells.cell[t].xf);
		add_property_long(cells,"ishiden",ws->rows.row[i].cells.cell[t].ishiden);
		add_property_long(cells,"width",ws->rows.row[i].cells.cell[t].width);
		add_property_long(cells,"colspan",ws->rows.row[i].cells.cell[t].colspan);
		add_property_long(cells,"rowspan",ws->rows.row[i].cells.cell[t].rowspan);
		add_property_long(cells,"l",ws->rows.row[i].cells.cell[t].l);
		add_property_double(cells,"d",ws->rows.row[i].cells.cell[t].d);
		
		if (ws->rows.row[i].cells.cell[t].str!=NULL)	add_property_string(cells,"str",ws->rows.row[i].cells.cell[t].str,1);

		add_index_zval(arr2,t,cells);
	}          
	add_property_zval(cell,"cell",arr2);	

//---------------------------------------------------------------------------------

		add_property_zval(rows,"cells",cell);
		add_index_zval(arr,i,rows);
	}          
	add_property_zval(row,"row",arr);	

	object_init(return_value);
	add_property_long(return_value,"defcolwidth",ws->defcolwidth);
	add_property_zval(return_value,"rows",row);
}


/* }}} */
/* The previous line is meant for vim and emacs, so it can correctly fold and 
   unfold functions in source code. See the corresponding marks just before 
   function definition, where the functions purpose is also documented. Please 
   follow this convention for the convenience of others editing your code.
*/


/*
 * Local variables:
 * tab-width: 4
 * c-basic-offset: 4
 * End:
 * vim600: noet sw=4 ts=4 fdm=marker
 * vim<600: noet sw=4 ts=4
 */

