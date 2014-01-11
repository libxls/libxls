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
 * Copyright 2014 David Hoerl
 *
 * USAGE: set the dump hander early on in your libxls code, before parsing a file: 
 *    xls_set_formula_hander(dump_formula);
 *
 */

#ifndef libXlsTester_xlsformula_h
#define libXlsTester_xlsformula_h

#ifdef AIX
#pragma pack(1)
#else
#pragma pack(push, 1)
#endif

#if defined(BLANK_CELL)

#include <libxls/xlstypes.h>
#include <libxls/xlsstruct.h>

#else 


// taken from the libxls files xlstypes.h and xlsstruct.h
#include <stdint.h>

typedef unsigned char		BYTE;
typedef uint16_t			WORD;
typedef uint32_t			DWORD;

#ifdef NO_ALIGN
typedef uint16_t			WORD_UA;
typedef uint32_t			DWORD_UA;
#else
typedef uint16_t			WORD_UA		__attribute__ ((aligned (1)));	// 2 bytes
typedef uint32_t			DWORD_UA	__attribute__ ((aligned (1)));	// 4 bytes
#endif

typedef struct FORMULA // BIFF8
{
    WORD	row;
    WORD	col;
    WORD	xf;
	// next 8 bytes either a IEEE double, or encoded on a byte basis
    BYTE	resid;
    BYTE	resdata[5];
    WORD	res;
    WORD	flags;
    BYTE	chn[4]; // BIFF8
    WORD	len;
    BYTE	value[1]; //var
}
FORMULA;

typedef struct FARRAY // BIFF8
{
    WORD	row1;
    WORD	row2;
    BYTE	col1;
    BYTE	col2;
    WORD	flags;
    BYTE	chn[4]; // BIFF8
    WORD	len;
    BYTE	value[1]; //var
}
FARRAY;
#endif

void dump_formula(WORD bof, WORD len, BYTE *formula);


#endif
