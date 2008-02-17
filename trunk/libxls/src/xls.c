#include <stdio.h>
#include <stdlib.h>
#include <memory.h>
#include <math.h>
#include <sys/types.h>
#include <string.h>
#include <wchar.h>

#include <libxls/xls.h>

#define min(a,b) ((a) < (b) ? (a) : (b))

//#define DEBUG

static char* libxls_version="1.0.0";

extern int xls(void)
{
    return 1;
}

void xls_addSST(xlsWorkBook* pWB,SST* sst,DWORD size)
{
    verbose("xls_addSST");
#ifdef DEBUG
    printf("xls_addSST\n");
#endif
    pWB->sst.continued=0;
    pWB->sst.lastln=0;
    pWB->sst.lastid=0;
    pWB->sst.lastrt=0;
    pWB->sst.lastsz=0;

    pWB->sst.count = sst->num;
    pWB->sst.string =(struct str_sst_string *) malloc (pWB->sst.count * sizeof(struct str_sst_string));
    xls_appendSST(pWB,&sst->strings,size-8);
}

void xls_appendSST(xlsWorkBook* pWB,BYTE* buf,DWORD size)
{
    DWORD ln; // String character count
    DWORD ofs; // Current offset in SST buffer
    DWORD rt; // Count or rich text formatting runs
    DWORD sz; // Size of asian phonetic settings block
    BYTE flag; // String flags

    char* ret;

    verbose("xls_appendSST\n");
#ifdef DEBUG
    printf("xls_appendSST %ld\n", size);
#endif

    ofs=0;

    while(ofs<size)
    {
        int ln_toread;

        // Restore state when we're in a continue record
        // or read string length
        if (pWB->sst.continued)
        {
            ln=pWB->sst.lastln;
            rt=pWB->sst.lastrt;
            sz=pWB->sst.lastsz;
        }
        else
        {
            ln=*(WORD*)(buf+ofs);
            rt = 0;
            sz = 0;

            ofs+=2;
        }

#ifdef DEBUG
        printf("ln=%ld\n", ln);
#endif

        // Read flags
        if (  (!pWB->sst.continued)
            ||(  (pWB->sst.continued)
               &&(ln != 0) ) )
        {
            flag=*(BYTE*)(buf+ofs);
            ofs++;

            // Count of rich text formatting runs
            if (flag & 0x8)
            {
                rt=*(WORD*)(buf+ofs);
                ofs+=2;
            }

            // Size of asian phonetic settings block
            if (flag & 0x4)
            {
                sz=*(DWORD*)(buf+ofs);
                ofs+=4;
#ifdef DEBUG
                printf("sz=%ld\n", sz);
#endif
            }
        }
        else
        {
            flag = 0;
        }

	// Read characters (compressed or not)
        ln_toread = 0;
        if (ln > 0)
        {
            if (flag & 0x1)
            {
                int new_len = 0;
                ln_toread = min((size-ofs)/2, ln);

                ret=utf8_decode(buf+ofs,ln_toread*2, &new_len,pWB->charset);
   
                if (ret == NULL)
                {
                    char *str = "*failed to decode utf16*";
                    ret = strdup(str);
                    new_len = strlen(str);
                }

                ret = (char *)realloc(ret,new_len+1);
                *(char*)(ret+new_len)=0;

                ln -= ln_toread;
                ofs+=ln_toread*2;
#ifdef DEBUG
                printf("String16: %s(%i)\n",ret,new_len);
#endif
            }
            else
            {
                ln_toread = min((size-ofs), ln);

                ret = (char *)malloc(ln_toread + 1);
                memcpy (ret, (buf+ofs),ln_toread);
                *(char*)(ret+ln_toread)=0;

                ln -= ln_toread;
                ofs+=ln_toread;
#ifdef DEBUG
                printf("String8: %s(%li) \n",ret,ln);
#endif
            }
        }
        else
        {
         ret = strdup("");
        }

        if (  (ln_toread > 0)
            ||(!pWB->sst.continued) )
        {
            // Concat string if it's a continue, or add string in table
            if (!pWB->sst.continued)
            {
                pWB->sst.lastid++;
                pWB->sst.string[pWB->sst.lastid-1].str=ret;
            }
            else
            {
                char *tmp;
                tmp=pWB->sst.string[pWB->sst.lastid-1].str;
                tmp=(char *)realloc(tmp,strlen(tmp)+strlen(ret)+1);
                pWB->sst.string[pWB->sst.lastid-1].str=tmp;
                memcpy(tmp+strlen(tmp),ret,strlen(ret)+1);
            }

#ifdef DEBUG
            printf("String % 4ld: %s<end>\n", pWB->sst.lastid-1, pWB->sst.string[pWB->sst.lastid-1].str);
#endif
        }

	// Jump list of rich text formatting runs
        if (  (ofs < size)
            &&(rt > 0) )
          {
           int rt_toread = min((size-ofs)/4, rt);
           rt -= rt_toread;
           ofs += rt_toread*4;
          }

	// Jump asian phonetic settings block
        if (  (ofs < size)
            &&(sz > 0) )
          {
           int sz_toread = min((size-ofs), sz);
           sz -= sz_toread;
           ofs += sz_toread;
          }

        pWB->sst.continued=0;
    }

    // Save current character count and count of rich text formatting runs and size of asian phonetic settings block
    if (ln > 0 || rt > 0 || sz > 0)
      {
       pWB->sst.continued = 1;
       pWB->sst.lastln = ln;
       pWB->sst.lastrt = rt;
       pWB->sst.lastsz = sz;

#ifdef DEBUG
       printf("continued: ln=%ld, rt=%ld, sz=%ld\n", ln, rt, sz);
#endif
     }
}

double NumFromRk(BYTE* rk)
{
    double num;
    DWORD drk;
    drk=*(DWORD*)rk;

    if(drk & 0x02)
    {
        num = (double)(drk >> 2);
    }
    else
    {
        *((DWORD *)&num+1) = drk & 0xfffffffc;
        *((DWORD *)&num) = 0;
    }
    if(drk & 0x01)
        num /= 100;
    return num;
}


void xls_addSheet(xlsWorkBook* pWB,BYTE* buf)
{
    BOUNDSHEET* tmp;

    verbose ("xls_addSheet");
    tmp=(BOUNDSHEET*)buf;

    if (pWB->sheets.count==0)
    {
        pWB->sheets.sheet=(struct st_sheet_data *) malloc((pWB->sheets.count+1)*sizeof (struct st_sheet_data));
    }
    else
    {
        pWB->sheets.sheet=(struct st_sheet_data *) realloc(pWB->sheets.sheet,(pWB->sheets.count+1)*sizeof (struct st_sheet_data));
    }
    pWB->sheets.sheet[pWB->sheets.count].name=get_unicode(&tmp->name,0);
    pWB->sheets.sheet[pWB->sheets.count].filepos=tmp->filepos;
    pWB->sheets.sheet[pWB->sheets.count].visibility=tmp->visibility;
    pWB->sheets.sheet[pWB->sheets.count].type=tmp->type;
    pWB->sheets.count++;
}


void xls_addRow(xlsWorkSheet* pWS,ROW* row)
{
    struct st_row_data* tmp;

    verbose ("xls_addRow");

    tmp=&pWS->rows.row[row->index];
    tmp->height=row->height;
    tmp->fcell=row->fcell;
    tmp->lcell=row->lcell;
    tmp->flags=row->flags;
    tmp->xf=row->xf&0xfff;
    tmp->xfflags=row->xf&0xf000;
    //	xls_showROW(tmp);
}

void xls_makeTable(xlsWorkSheet* pWS)
{
    WORD i,t;
    struct st_row_data* tmp;
    verbose ("xls_makeTable");

    pWS->rows.row=(struct st_row_data *) malloc((pWS->rows.lastrow+1)*sizeof(struct st_row_data));

    for (t=0;t<=pWS->rows.lastrow;t++)
    {
        tmp=&pWS->rows.row[t];
        tmp->index=t;
        tmp->fcell=0;
        tmp->lcell=pWS->rows.lastcol;

        tmp->cells.cell=(struct st_cell_data *) malloc((pWS->rows.lastcol+1)*sizeof(struct st_cell_data));

        for (i=0;i<=pWS->rows.lastcol;i++)
        {
            tmp->cells.cell[i].col=i;
            tmp->cells.cell[i].row=t;
            tmp->cells.cell[i].width=pWS->defcolwidth;
            tmp->cells.cell[i].xf=0;
            tmp->cells.cell[i].str=NULL;
            tmp->cells.cell[i].d=0;
            tmp->cells.cell[i].l=0;
            tmp->cells.cell[i].ishiden=0;
            tmp->cells.cell[i].colspan=0;
            tmp->cells.cell[i].rowspan=0;
            tmp->cells.cell[i].id=0x201;
            tmp->cells.cell[i].str=NULL;
        }
    }
}


void xls_addCell(xlsWorkSheet* pWS,BOF* bof,BYTE* buf)
{
    struct st_cell_data*	cell;
    struct st_row_data*	row;
    int i;

    verbose ("xls_addCell");

    row=&pWS->rows.row[((COL*)buf)->row];
    cell=&row->cells.cell[((COL*)buf)->col];

    cell->id=bof->id;
    cell->xf=((COL*)buf)->xf;

    switch (bof->id)
    {
    case 0x06:	//FORMULA
        if (((FORMULA*)buf)->res!=0xffff)
            cell->d=*(double *)&((FORMULA*)buf)->resid;
        cell->str=xls_getfcell(pWS->workbook,cell);
        break;
    case 0x0BD:	//MULRK
        for (i=0;i<=*(WORD *)(buf+(bof->size-2))-((COL*)buf)->col;i++)
        {
            cell=&row->cells.cell[((COL*)buf)->col+i];
            //				col=row->cols[i];
            cell->id=bof->id;
            cell->xf=*((WORD *)(buf+(4+i*6)));
            cell->d=NumFromRk((BYTE *)(buf+(4+i*6+2)));
            cell->str=xls_getfcell(pWS->workbook,cell);
        }
        break;
    case 0x0BE:	//MULBLANK
        for (i=0;i<=*(WORD *)(buf+(bof->size-2))-((COL*)buf)->col;i++)
        {
            cell=&row->cells.cell[((COL*)buf)->col+i];
            //				col=row->cols[i];
            cell->id=bof->id;
            cell->xf=*((WORD *)(buf+(4+i*2)));
            cell->str=xls_getfcell(pWS->workbook,cell);
        }
        break;
    case 0xFD:	//LABELSST
        cell->l=*(DWORD *)&((LABELSST*)buf)->value;
        cell->str=xls_getfcell(pWS->workbook,cell);
        break;
    case 0x27E:	//RK
        cell->d=NumFromRk(((RK*)buf)->value);
        cell->str=xls_getfcell(pWS->workbook,cell);
        break;
    case 0x203:	//NUMBER
        cell->d=*(double *)&((BR_NUMBER*)buf)->value;
        cell->str=xls_getfcell(pWS->workbook,cell);
        break;
    default:
        cell->str=xls_getfcell(pWS->workbook,cell);
        break;
    }
    // xls_showCell(cell);
}

void xls_addFont(xlsWorkBook* pWB,FONT* font)
{
    struct st_font_data* tmp;

    verbose("xls_addFont");
    if (pWB->fonts.count==0)
    {
        pWB->fonts.font=(struct st_font_data *) malloc(sizeof(struct st_font_data));
    }
    else
    {
        pWB->fonts.font=(struct st_font_data *) realloc(pWB->fonts.font,(pWB->fonts.count+1)*sizeof(struct st_font_data));
    }

    tmp=&pWB->fonts.font[pWB->fonts.count];

    tmp->name=get_unicode((BYTE*)&font->name,0);
    tmp->height=font->height;
    tmp->flag=font->flag;
    tmp->color=font->color;
    tmp->bold=font->bold;
    tmp->escapement=font->escapement;
    tmp->underline=font->underline;
    tmp->family=font->family;
    tmp->charset=font->charset;

    //	xls_showFont(tmp);
    pWB->fonts.count++;
}

void xls_addFormat(xlsWorkBook* pWB,FORMAT* format)
{
    struct st_format_data* tmp;

    verbose("xls_addFormat");
    if (pWB->formats.count==0)
    {
        pWB->formats.format=(struct st_format_data *) malloc(sizeof(struct st_format_data));
    }
    else
    {
        pWB->formats.format=(struct st_format_data *) realloc(pWB->formats.format,(pWB->formats.count+1)*sizeof(struct st_format_data));
    }

    tmp=&pWB->formats.format[pWB->formats.count];

    tmp->index=format->index;
    tmp->value=get_unicode((BYTE*)format + 2,1);

    //	xls_showFormat(tmp);
    pWB->formats.count++;
}

void xls_addXF(xlsWorkBook* pWB,XF* xf)
{
    struct st_xf_data* tmp;

    verbose("xls_addXF");
    if (pWB->xfs.count==0)
    {
        pWB->xfs.xf=(struct st_xf_data *) malloc(sizeof(struct st_xf_data));
    }
    else
    {
        pWB->xfs.xf=(struct st_xf_data *) realloc(pWB->xfs.xf,(pWB->xfs.count+1)*sizeof(struct st_xf_data));
    }

    tmp=&pWB->xfs.xf[pWB->xfs.count];

    tmp->font=xf->font;
    tmp->format=xf->format;
    tmp->type=xf->type;
    tmp->align=xf->align;
    tmp->rotation=xf->rotation;
    tmp->ident=xf->ident;
    tmp->usedattr=xf->usedattr;
    tmp->linestyle=xf->linestyle;
    tmp->linecolor=xf->linecolor;
    tmp->groundcolor=xf->groundcolor;

    //	xls_showXF(tmp);
    pWB->xfs.count++;
}

void xls_addColinfo(xlsWorkSheet* pWS,COLINFO* colinfo)
{
    struct st_colinfo_data* tmp;

    verbose("xls_addColinfo");
    if (pWS->colinfo.count==0)
    {
        pWS->colinfo.col=(struct st_colinfo_data *) malloc(sizeof(struct st_colinfo_data));
    }
    else
    {
        pWS->colinfo.col=(struct st_colinfo_data *) realloc(pWS->colinfo.col,(pWS->colinfo.count+1)*sizeof(struct st_colinfo_data));
    }

    tmp=&pWS->colinfo.col[pWS->colinfo.count];
    tmp->first=colinfo->first;
    tmp->last=colinfo->last;
    tmp->width=colinfo->width;
    tmp->xf=colinfo->xf;
    tmp->flags=colinfo->flags;

    //	xls_showColinfo(tmp);
    pWS->colinfo.count++;
}

void xls_mergedCells(xlsWorkSheet* pWS,BOF* bof,BYTE* buf)
{
    int count=*((WORD*)buf);
    int i,c,r;
    struct MERGEDCELLS* span;
    verbose("Merged Cells");
    for (i=0;i<count;i++)
    {
        span=(struct MERGEDCELLS*)(buf+(2+i*sizeof(struct MERGEDCELLS)));
        //		printf("Merged Cells: [%i,%i] [%i,%i] \n",span->colf,span->rowf,span->coll,span->rowl);
        for (r=span->rowf;r<=span->rowl;r++)
            for (c=span->colf;c<=span->coll;c++)
                pWS->rows.row[r].cells.cell[c].ishiden=1;
        pWS->rows.row[span->rowf].cells.cell[span->colf].colspan=(span->coll-span->colf+1);
        pWS->rows.row[span->rowf].cells.cell[span->colf].rowspan=(span->rowl-span->rowf+1);
        pWS->rows.row[span->rowf].cells.cell[span->colf].ishiden=0;
    }
}

extern void xls_parseWorkBook(xlsWorkBook* pWB)
{
    BOF bof1;
    BOF bof2;
    BYTE* buf;
    DWORD size;
    //FILE* file;
    verbose ("xls_ParseWorkBook");
    //file=pWB->file;

    do
    {
        ole2_read(&bof1, 1,4,pWB->olest);
        size=bof1.size;
        buf=(BYTE *)malloc(size);
        ole2_read(buf, 1,size,pWB->olest);
        //	xls_showBOF(&bof1);
        switch (bof1.id)
        {
        case 0x0A:	//EOF
            free(buf);
            break;
        case 0x809:	//BIFF5-8
            if (((BIFF*)buf)->ver==0x600)
                pWB->is5ver=0;
            else
                pWB->is5ver=1;
            pWB->type=((BIFF*)buf)->type;
            free(buf);
            break;
        case 0x042: //CODEPAGE
            pWB->codepage=*(WORD*)buf;
            free(buf);
            break;
        case 0x3c: //CONTINUE
            if (bof2.id==0xfc)
                xls_appendSST(pWB,buf,bof1.size);
            bof1=bof2;
            free(buf);
            break;
        case 0xfc: //SST
            xls_addSST(pWB,(SST*)buf,bof1.size);
            free(buf);
            break;
        case 0xff: //EXTSST
            //		dumpbuf("EXTSST",bof1.size,buf);
            free(buf);
            break;
        case 0x85:	//BOUNDSHEET
            xls_addSheet(pWB,buf);
            free(buf);
            break;
        case 0x0e0:  	//XF
            xls_addXF(pWB,(XF*)buf);
            free(buf);
            break;
        case 0x031:  	//FONT
            xls_addFont(pWB,(FONT*)buf);
            free(buf);
            break;
        case 0x41E:	//FORMAT
            xls_addFormat(pWB,(FORMAT*)buf);
            free(buf);
            break;
        default:
            free(buf);
            break;
        }

        bof2=bof1;

    }
    while ((!pWB->olest->eof)&&(bof1.id!=0x0A));
}


void xls_preparseWorkSheet(xlsWorkSheet* pWS)
{
    BOF tmp;
    BYTE* buf;
    verbose ("xls_PreParseWorkSheet");
    ole2_seek(pWS->workbook->olest,pWS->filepos);
    do
    {
        ole2_read(&tmp, 1,4,pWS->workbook->olest);
        buf=(BYTE *)malloc(tmp.size);
        ole2_read(buf, 1,tmp.size,pWS->workbook->olest);
        //	xls_showBOF(&tmp);
        switch (tmp.id)
        {
        case 0x55:     //DEFCOLWIDTH
            pWS->defcolwidth=*(WORD*)buf*256;
            break;
        case 0x7D:     //COLINFO
            xls_addColinfo(pWS,(COLINFO*)buf);
            break;
        case 0x208:	//ROW
            if (pWS->rows.lastcol<((ROW*)buf)->lcell-1)
                pWS->rows.lastcol=((ROW*)buf)->lcell-1;
            if (pWS->rows.lastrow<((ROW*)buf)->index)
                pWS->rows.lastrow=((ROW*)buf)->index;
            break;
        }
        free(buf);
    }
    while ((!pWS->workbook->olest->eof)&&(tmp.id!=0x0A));
}

void xls_formatColumn(xlsWorkSheet* pWS)
{
    int i,t,ii;
    int fcol,lcol;

    for (i=0;i<pWS->colinfo.count;i++)
    {
        if (pWS->colinfo.col[i].first<=pWS->rows.lastcol)
            fcol=pWS->colinfo.col[i].first;
        else
            fcol=pWS->rows.lastcol;

        if (pWS->colinfo.col[i].last<=pWS->rows.lastcol)
            lcol=pWS->colinfo.col[i].last;
        else
            lcol=pWS->rows.lastcol;

        for (t=fcol;t<=lcol;t++)
            for (ii=0;ii<=pWS->rows.lastrow;ii++)
            {
                if (pWS->colinfo.col[i].flags&1)
                    pWS->rows.row[ii].cells.cell[t].ishiden=1;
                pWS->rows.row[ii].cells.cell[t].width=pWS->colinfo.col[i].width;
            }
    }
}

extern void xls_parseWorkSheet(xlsWorkSheet* pWS)
{
    BOF tmp;
    BYTE* buf;
    verbose ("xls_ParseWorkSheet");

    xls_preparseWorkSheet(pWS);
    xls_makeTable(pWS);
    xls_formatColumn(pWS);

    ole2_seek(pWS->workbook->olest,pWS->filepos);
    do
    {
        ole2_read(&tmp, 1,4,pWS->workbook->olest);
        buf=(BYTE *)malloc(tmp.size);
        ole2_read(buf, 1,tmp.size,pWS->workbook->olest);

        //	xls_showBOF(&tmp);
        switch (tmp.id)
        {
        case 0x0A:	//EOF
            break;
        case 0x0E5: 	//MERGEDCELLS
            xls_mergedCells(pWS,&tmp,buf);
            break;
        case 0x208:	//ROW
            xls_addRow(pWS,(ROW*)buf);
            break;
        case 0x20B:	//INDEX
            break;
        case 0x0BD:	//MULRK
        case 0x0BE:	//MULBLANK
        case 0x203:	//NUMBER
        case 0x06:		//FORMULA
        case 0x27e:	//RK
        case 0xFD: 	//LABELSST
        case 0x201:	//BLANK
            xls_addCell(pWS,&tmp,buf);
            break;
        default:
            break;
        }

        free(buf);
    }
    while ((!pWS->workbook->olest->eof)&&(tmp.id!=0x0A));


}

extern xlsWorkSheet * xls_getWorkSheet(xlsWorkBook* pWB,int num)
{
    xlsWorkSheet * pWS;
    verbose ("xls_getWorkSheet");
    pWS=(xlsWorkSheet *) malloc(sizeof(xlsWorkSheet));
    pWS->filepos=pWB->sheets.sheet[num].filepos;
    pWS->workbook=pWB;
    pWS->rows.lastcol=0;
    pWS->rows.lastrow=0;
    pWS->colinfo.count=0;
    return(pWS);
}

extern xlsWorkBook* xls_open(char *file,char* charset)
{
    xlsWorkBook* pWB;
    OLE2*		ole;

    pWB=(xlsWorkBook*)malloc(sizeof(xlsWorkBook));
    verbose ("xls_open");

    // open excel file
    if (!(ole=ole2_open(file)))
    {
        printf("Cannot open File \"%s\"",file);
        return(NULL);
    }

    // open Workbook
    if (!(pWB->olest=ole2_fopen(ole,"Workbook")))
    {
        printf("Workbook not found\n");
        return(NULL);
    }

    pWB->sheets.count=0;
    pWB->xfs.count=0;
    pWB->fonts.count=0;
    pWB->charset = (char *)malloc(strlen(charset) * sizeof(char));
    strcpy(pWB->charset,charset);
    xls_parseWorkBook(pWB);
    return(pWB);
}

extern void xls_close(xlsWorkBook* pWB)
{
	OLE2*		ole;
	verbose ("xls_close");
	ole=pWB->olest->ole;
	ole2_fclose(pWB->olest);
	ole2_close(ole);
	free(pWB);
}

extern const char* xls_getVersion(void)
{
    return libxls_version;
}

