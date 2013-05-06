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

#include <libxls/xlsstruct.h>

int is_bigendian();
int intVal (int i);
short shortVal (short s);

void convertHeader(OLE2Header *h);
void convertPss(PSS* pss);

void convertDouble(BYTE *d);
void convertBof(BOF *b);
void convertBiff(BIFF *b);
void convertWindow(WIND1 *w);
void convertSst(SST *s);
void convertXf5(XF5 *x);
void convertXf8(XF8 *x);
void convertFont(FONT *f);
void convertFormat(FORMAT *f);
void convertBoundsheet(BOUNDSHEET *b);
void convertColinfo(COLINFO *c);
void convertRow(ROW *r);
void convertMergedcells(MERGEDCELLS *m);
void convertCol(COL *c);
void convertFormula(FORMULA *f);
void convertHeader(OLE2Header *h);
void convertPss(PSS* pss);
void convertUnicode(wchar_t *w, char *s, int len);

#define W_ENDIAN(a) a=shortVal(a)
#define D_ENDIAN(a) a=intVal(a)
