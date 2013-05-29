/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * This file is part of libxls -- A multiplatform, C/C++ library
 * for parsing Excel(TM) files.
 *
 * Redistribution and use in source and binary forms, with or without modification, are
 * permitted provided that the following conditions are met:
 *
 *    1. Redistributions of source code must retain the above copyright notice, this list of
 *       conditions and the following disclaimer.
 *
 *    2. Redistributions in binary form must reproduce the above copyright notice, this list
 *       of conditions and the following disclaimer in the documentation and/or other materials
 *       provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY David Hoerl ''AS IS'' AND ANY EXPRESS OR IMPLIED
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
 * FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL David Hoerl OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 * SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 * ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 * Copyright 2013 Bob Colbert
 *
 */

#include <stdlib.h>

#include "libxls/xlstypes.h"
#include "libxls/endian.h"
#include "libxls/ole.h"

int is_bigendian()
{
#if defined (__BIG_ENDIAN__)
    return 1;
#elif defined (__LITTLE_ENDIAN__)
    return 0;
#else
#warning NO ENDIAN
    static int n = 1;

    if (*(char *)&n == 1)
    {
        return 0;
    }
    else
    {
        return 1;
    }
#endif
}

int intVal (int i)
{
    unsigned char c1, c2, c3, c4;

    if (is_bigendian()) {
        c1 = i & 255;
        c2 = (i >> 8) & 255;
        c3 = (i >> 16) & 255;
        c4 = (i >> 24) & 255;

        return ((int)c1 << 24) + ((int)c2 << 16) + ((int)c3 << 8) + c4;
    } else {
        return i;
    }
}

short shortVal (short s)
{
    unsigned char c1, c2;
    
    if (is_bigendian()) {
        c1 = s & 255;
        c2 = (s >> 8) & 255;
    
        return (c1 << 8) + c2;
    } else {
        return s;
    }
}

void convertDouble(unsigned char *d)
{
    unsigned char t;
    int i;

    if (is_bigendian()) {
        for (i=0; i<4; i++)
        {
            t = d[7-i];
            d[8-i] = d[i];
            d[i] = t;
        }
    }
}

void convertBof(BOF *b)
{
    b->id = shortVal(b->id);
    b->size = shortVal(b->size);
}

void convertBiff(BIFF *b)
{
    b->ver = shortVal(b->ver);
    b->type = shortVal(b->type);
    b->id_make = shortVal(b->id_make);
    b->year = shortVal(b->year);
    b->flags = intVal(b->flags);
    b->min_ver = intVal(b->min_ver);
}

void convertWindow(WIND1 *w)
{
    w->xWn = shortVal(w->xWn);
    w->yWn = shortVal(w->yWn);
    w->dxWn = shortVal(w->dxWn);
    w->dyWn = shortVal(w->dyWn);
    w->grbit = shortVal(w->grbit);
    w->itabCur = shortVal(w->itabCur);
    w->itabFirst = shortVal(w->itabFirst);
    w->ctabSel = shortVal(w->ctabSel);
    w->wTabRatio = shortVal(w->wTabRatio);
}

void convertSst(SST *s)
{
    s->num=intVal(s->num);
    s->num=intVal(s->numofstr);
}

void convertXf5(XF5 *x)
{
    x->font=shortVal(x->font);
    x->format=shortVal(x->format);
    x->type=shortVal(x->type);
    x->align=shortVal(x->align);
    x->color=shortVal(x->color);
    x->fill=shortVal(x->fill);
    x->border=shortVal(x->border);
    x->linestyle=shortVal(x->linestyle);
}

void convertXf8(XF8 *x)
{
    W_ENDIAN(x->font);
    W_ENDIAN(x->format);
    W_ENDIAN(x->type);
    D_ENDIAN(x->linestyle);
    D_ENDIAN(x->linecolor);
    W_ENDIAN(x->groundcolor);
}

void convertFont(FONT *f)
{
    W_ENDIAN(f->height);
    W_ENDIAN(f->flag);
    W_ENDIAN(f->color);
    W_ENDIAN(f->bold);
    W_ENDIAN(f->escapement);
}

void convertFormat(FORMAT *f)
{
    W_ENDIAN(f->index);
}

void convertBoundsheet(BOUNDSHEET *b)
{
    D_ENDIAN(b->filepos);
}

void convertColinfo(COLINFO *c)
{
    W_ENDIAN(c->first);
    W_ENDIAN(c->last);
    W_ENDIAN(c->width);
    W_ENDIAN(c->xf);
    W_ENDIAN(c->flags);
    W_ENDIAN(c->notused);
}

void convertRow(ROW *r)
{
    W_ENDIAN(r->index);
    W_ENDIAN(r->fcell);
    W_ENDIAN(r->lcell);
    W_ENDIAN(r->height);
    W_ENDIAN(r->notused);
    W_ENDIAN(r->notused2);
    W_ENDIAN(r->flags);
    W_ENDIAN(r->xf);
}

void convertMergedcells(MERGEDCELLS *m)
{
    W_ENDIAN(m->rowf);
    W_ENDIAN(m->rowl);
    W_ENDIAN(m->colf);
    W_ENDIAN(m->coll);
}

void convertCol(COL *c)
{
    W_ENDIAN(c->row);
    W_ENDIAN(c->col);
    W_ENDIAN(c->xf);
}

void convertFormula(FORMULA *f)
{
    W_ENDIAN(f->row);
    W_ENDIAN(f->col);
    W_ENDIAN(f->xf);
    convertDouble((BYTE *)&f->resid);
/*
    D_ENDIAN(f->res);
*/
    W_ENDIAN(f->flags);
    W_ENDIAN(f->len);
    fflush(stdout);
}

void convertHeader(OLE2Header *h)
{
    int i;
    for (i=0; i<2; i++)
        h->id[i] = intVal(h->id[i]);
    for (i=0; i<4; i++)
        h->clid[i] = intVal(h->clid[i]);
    h->verminor  = shortVal(h->verminor);
    h->verdll    = shortVal(h->verdll);
    h->byteorder = shortVal(h->byteorder);
    h->lsectorB  = shortVal(h->lsectorB);
    h->lssectorB = shortVal(h->lssectorB);
    h->reserved1 = shortVal(h->reserved1);
    h->reserved2 = intVal(h->reserved2);
    h->reserved3 = intVal(h->reserved3);

    h->cfat      = intVal(h->cfat);
    h->dirstart  = intVal(h->dirstart);

    h->reserved4 = intVal(h->reserved4);

    h->sectorcutoff = intVal(h->sectorcutoff);
    h->sfatstart = intVal(h->sfatstart);
    h->csfat = intVal(h->csfat);
    h->difstart = intVal(h->difstart);
    h->cdif = intVal(h->cdif);
    for (i=0; i<109; i++)
        h->MSAT[i] = intVal(h->MSAT[i]);
}

void convertPss(PSS* pss)
{
    int i;
    pss->bsize = shortVal(pss->bsize);
    pss->left  = intVal(pss->left);
    pss->right  = intVal(pss->right);
    pss->child  = intVal(pss->child);

    for(i=0; i<8; i++)
        pss->guid[i]=shortVal(pss->guid[i]);
    pss->userflags  = intVal(pss->userflags);
/*    TIME_T	time[2]; */
    pss->sstart  = intVal(pss->sstart);
    pss->size  = intVal(pss->size);
    pss->proptype  = intVal(pss->proptype);
}

void convertUnicode(wchar_t *w, char *s, int len)
{
    short *x;
    int i;

    x=(short *)s;
    w = (wchar_t*)malloc((len+1)*sizeof(wchar_t));

    for(i=0; i<len; i++)
    {
        w[i]=shortVal(x[i]);
    }
    w[len] = '\0';
}
