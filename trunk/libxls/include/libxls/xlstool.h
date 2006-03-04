#include <libxls/xlsstruct.h>

static const int colors[] =
    {
        0x000000,
        0xFFFFFF,
        0xFF0000,
        0x00FF00,
        0x0000FF,
        0xFFFF00,
        0xFF00FF,
        0x00FFFF,
        0x800000,
        0x008000,
        0x000080,
        0x808000,
        0x800080,
        0x008080,
        0xC0C0C0,
        0x808080,
        0x9999FF,
        0x993366,
        0xFFFFCC,
        0xCCFFFF,
        0x660066,
        0xFF8080,
        0x0066CC,
        0xCCCCFF,
        0x000080,
        0xFF00FF,
        0xFFFF00,
        0x00FFFF,
        0x800080,
        0x800000,
        0x008080,
        0x0000FF,
        0x00CCFF,
        0xCCFFFF,
        0xCCFFCC,
        0xFFFF99,
        0x99CCFF,
        0xFF99CC,
        0xCC99FF,
        0xFFCC99,
        0x3366FF,
        0x33CCCC,
        0x99CC00,
        0xFFCC00,
        0xFF9900,
        0xFF6600,
        0x666699,
        0x969696,
        0x003366,
        0x339966,
        0x003300,
        0x333300,
        0x993300,
        0x993366,
        0x333399,
        0x333333
    };

void dumpbuf(char* fname,long size,BYTE* buf);
void verbose(char* str);
char* utf8_decode(const char *s, int len, int *newlen, const char* encoding);
char*  get_unicode(BYTE *s,BYTE is2);
DWORD xls_getColor(const WORD color,WORD def);

extern void xls_showBookInfo(xlsWorkBook* pWB);
extern void xls_showROW(struct st_row_data* row);
extern void xls_showColinfo(struct st_colinfo_data* col);
extern void xls_showCell(struct st_cell_data* cell);
extern void xls_showFont(struct st_font_data* font);
extern void xls_showFormat(struct st_format_data* format);
extern void xls_showXF(struct st_xf_data* xf);
extern char* xls_getfcell(xlsWorkBook* pWB,struct st_cell_data* cell);
extern char* xls_getCSS(xlsWorkBook* pWB);
