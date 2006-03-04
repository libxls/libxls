#include <libxls/xlstool.h>

extern int xls(void);
void xls_addSST(xlsWorkBook* pWB,SST* sst,DWORD size);
void xls_appendSST(xlsWorkBook* pWB,BYTE* buf,DWORD size);

extern void xls_parseWorkBook(xlsWorkBook* pWB);
extern void xls_parseWorkSheet(xlsWorkSheet* pWS);
extern xlsWorkSheet * xls_getWorkSheet(xlsWorkBook* pWB,int num);
extern xlsWorkBook* xls_open(char *file,char* charset);
extern void xls_close(xlsWorkBook* pWB);
