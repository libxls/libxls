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
 */
#include <stdio.h>
#include <assert.h>

#include "xlsformula.h"

//#include <libxls/xls.h>

static void dump_formula_formula(WORD len, BYTE *buf);
static void dump_formula_array(WORD len, BYTE *buf);
static void dump_formula_data(WORD len, BYTE *buf);
static WORD func_len(WORD func);
static WORD get_token_size(BYTE *buf);

typedef struct { const char *xlsName; const char *excelName; } excelNames;
static const excelNames tokenNames[128];
static const excelNames functionNames[368];

static unsigned short xlsShortVal (short s);

void dump_formula(WORD bof, WORD len, BYTE *buf)
{
	if(bof == 0x0221) {
		dump_formula_array(len, buf);
	} else {
		dump_formula_formula(len, buf);
	}
}

static void dump_formula_formula(WORD len, BYTE *buf)
{
	FORMULA *f = (FORMULA *)buf;

	printf("FORMULA LEN: %d\n", len);
	printf("ROW: %d\n", f->row);
	printf("COL: %d\n", f->col);
	printf("NUM: "); for(int i=-1; i<7; ++i) printf("%2.2x ", f->resdata[i]); printf("\n");
	printf("OPTIONS: 0x%x\n", f->flags);
	dump_formula_data(f->len, f->value);
}

static void dump_formula_array(WORD len, BYTE *buf)
{
	FARRAY *f = (FARRAY *)buf;

	printf("FORMULA ARRAY LEN: %d\n", len);
	printf("ROW: %d->%d\n", f->row1, f->row2);
	printf("COL: %d->%d\n", f->col1, f->col2);
	printf("OPTIONS: 0x%x\n", f->flags);
	dump_formula_data(f->len, f->value);
}

static void dump_formula_data(WORD flen, BYTE *buf)
{
	printf("FORMULA LEN: %d\n", flen);
	if(flen) {
#if 0
		printf("   ");
		for(int i=0; i<f->len; ++i) printf("%2.2x ", buf[i]);
		printf("\n");
#endif
		BYTE *b = buf;
		while((b-buf) < flen) {
			WORD len = get_token_size(b);
			if(!len) {
				printf("YIKES: token 0x%2.2x no len!\n", b[0]);
				return;
			}
			
			int printBytes = 1;
			excelNames tn = tokenNames[ b[0] ];
			switch(b[0]) {
			case 0x21:
			case 0x41:
			case 0x61:
			{
				unsigned short func = xlsShortVal(*(short *)&b[1]);
				excelNames fn = functionNames[ func ];
				printf("%s (0x%x) %s", tn.xlsName, b[0], fn.xlsName);
				printBytes = 0;
			}	break;
			case 0x22:
			case 0x42:
			case 0x62:
			{
				unsigned short func = xlsShortVal(*(short *)&b[2]);
				excelNames fn = functionNames[ func ];
				printf("%s (0x%x) %s arguments=%d", tn.xlsName, b[0], fn.xlsName, b[1]);
				printBytes = 0;
			}	break;

			case 0x24:
			case 0x44:
			case 0x64:
			{
				unsigned short col = xlsShortVal(*(short *)&b[3]);
				unsigned short flags = col & 0xC000;
				col &= ~0xC000;
				printf("%s (0x%x) ROW=%d COL=%d FLAGS=0x%x", tn.xlsName, b[0], xlsShortVal(*(short *)&b[1]), col, flags );
				printBytes = 0;
			}	break;
			
			default:
				//printf("TOKEN[0x%x]: ", b[0]);
				printf("%s (0x%x): ", tn.xlsName, b[0]);
				break;
			}
			if(printBytes) {
				for(int i=1; i<len; ++i) printf("%2.2x ", b[i]);
			}
			printf("\n");
			b += len;
		}
	}
}



// From "Open Office MS Excel File Format", pg 42-43
static WORD token_size[128] = {
	0,
	5,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	1,
	0,
	0,
	0,
	0,
	0,
	2,
	2,
	3,
	9,
	
	8,
	3,
	4,
	5,
	5,
	9,
	7,
	7,
	7,
	3,
	5,
	9,
	5,
	9,
	3,
	3,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	7,
	7,
	11,
	7,
	11,
	0,
	0,

	8,
	3,
	4,
	5,
	5,
	9,
	7,
	7,
	7,
	3,
	5,
	9,
	5,
	9,
	3,
	3,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	7,
	7,
	11,
	7,
	11,
	0,
	0,

	8,
	3,
	4,
	5,
	5,
	9,
	7,
	7,
	7,
	3,
	5,
	9,
	5,
	9,
	3,
	3,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	0,
	7,
	7,
	11,
	7,
	11,
	0,
	0,
};

static WORD get_token_size(BYTE *buf)
{
	WORD len = token_size[buf[0]];
	if(!len) {
		switch(buf[0]) {
		case 0x19:
		{
			BYTE flag = buf[1];
			switch(flag) {
			case 0x00:	// only in original OSX Excel circa 2001
			case 0x01:
			case 0x02:
				len = 4;
				break;
			case 0x04:	// not handled
				break;
			case 0x08:
			case 0x10:
			case 0x20:
			case 0x40:
			case 0x41:
				len = 4;
				break;
			}
			if(!len) printf("YIKES: 0x19 with flag 0x%2.2x\n", flag);
		}	break;
		default:
			break;
		}
	}
	return len;
}

static const excelNames tokenNames[128] = {
    { "",               "" },               // 0x00 (0)
    { "OP_EXP",         "ptgExp" },         // 0x01 (1)
    { "OP_TBL",         "ptgTbl" },         // 0x02 (2)
    { "OP_ADD",         "ptgAdd" },         // 0x03 (3)
    { "OP_SUB",         "ptgSub" },         // 0x04 (4)
    { "OP_MUL",         "ptgMul" },         // 0x05 (5)
    { "OP_DIV",         "ptgDiv" },         // 0x06 (6)
    { "OP_POWER",       "ptgPower" },       // 0x07 (7)
    { "OP_CONCAT",      "ptgConcat" },      // 0x08 (8)
    { "OP_LT",          "ptgLT" },          // 0x09 (9)
    { "OP_LE",          "ptgLE" },          // 0x0a (10)
    { "OP_EQ",          "ptgEQ" },          // 0x0b (11)
    { "OP_GE",          "ptgGE" },          // 0x0c (12)
    { "OP_GT",          "ptgGT" },          // 0x0d (13)
    { "OP_NE",          "ptgNE" },          // 0x0e (14)
    { "OP_ISECT",       "ptgIsect" },       // 0x0f (15)
    { "OP_UNION",       "ptgUnion" },       // 0x10 (16)
    { "OP_RANGE",       "ptgRange" },       // 0x11 (17)
    { "OP_UPLUS",       "ptgUplus" },       // 0x12 (18)
    { "OP_UMINUS",      "ptgUminus" },      // 0x13 (19)
    { "OP_PERCENT",     "ptgPercent" },     // 0x14 (20)
    { "OP_PAREN",       "ptgParen" },       // 0x15 (21)
    { "OP_MISSARG",     "ptgMissArg" },     // 0x16 (22)
    { "OP_STR",         "ptgStr" },         // 0x17 (23)
    { "",               "" },               // 0x18 (24)
    { "OP_ATTR",        "ptgAttr" },        // 0x19 (25)
    { "OP_SHEET",       "ptgSheet" },       // 0x1a (26)
    { "OP_ENDSHEET",    "ptgEndSheet" },    // 0x1b (27)
    { "OP_ERR",         "ptgErr" },         // 0x1c (28)
    { "OP_BOOL",        "ptgBool" },        // 0x1d (29)
    { "OP_INT",         "ptgInt" },         // 0x1e (30)
    { "OP_NUM",         "ptgNum" },         // 0x1f (31)
    { "OP_ARRAY",       "ptgArray" },       // 0x20 (32)
    { "OP_FUNC",        "ptgFunc" },        // 0x21 (33)
    { "OP_FUNCVAR",     "ptgFuncVar" },     // 0x22 (34)
    { "OP_NAME",        "ptgName" },        // 0x23 (35)
    { "OP_REF",         "ptgRef" },         // 0x24 (36)
    { "OP_AREA",        "ptgArea" },        // 0x25 (37)
    { "OP_MEMAREA",     "ptgMemArea" },     // 0x26 (38)
    { "OP_MEMERR",      "ptgMemErr" },      // 0x27 (39)
    { "OP_MEMNOMEM",    "ptgMemNoMem" },    // 0x28 (40)
    { "OP_MEMFUNC",     "ptgMemFunc" },     // 0x29 (41)
    { "OP_REFERR",      "ptgRefErr" },      // 0x2a (42)
    { "OP_AREAERR",     "ptgAreaErr" },     // 0x2b (43)
    { "OP_REFN",        "ptgRefN" },        // 0x2c (44)
    { "OP_AREAN",       "ptgAreaN" },       // 0x2d (45)
    { "OP_MEMAREAN",    "ptgMemAreaN" },    // 0x2e (46)
    { "OP_MEMNOMEMN",   "ptgMemNoMemN" },   // 0x2f (47)
    { "",               "" },               // 0x30 (48)
    { "",               "" },               // 0x31 (49)
    { "",               "" },               // 0x32 (50)
    { "",               "" },               // 0x33 (51)
    { "",               "" },               // 0x34 (52)
    { "",               "" },               // 0x35 (53)
    { "",               "" },               // 0x36 (54)
    { "",               "" },               // 0x37 (55)
    { "",               "" },               // 0x38 (56)
    { "OP_NAMEX",       "ptgNameX" },       // 0x39 (57)
    { "OP_REF3D",       "ptgRef3d" },       // 0x3a (58)
    { "OP_AREA3D",      "ptgArea3d" },      // 0x3b (59)
    { "OP_REFERR3D",    "ptgRefErr3d" },    // 0x3c (60)
    { "OP_AREAERR3D",   "ptgAreaErr3d" },   // 0x3d (61)
    { "",               "" },               // 0x3e (62)
    { "",               "" },               // 0x3f (63)
    { "OP_ARRAYV",      "ptgArrayV" },      // 0x40 (64)
    { "OP_FUNCV",       "ptgFuncV" },       // 0x41 (65)
    { "OP_FUNCVARV",    "ptgFuncVarV" },    // 0x42 (66)
    { "OP_NAMEV",       "ptgNameV" },       // 0x43 (67)
    { "OP_REFV",        "ptgRefV" },        // 0x44 (68)
    { "OP_AREAV",       "ptgAreaV" },       // 0x45 (69)
    { "OP_MEMAREAV",    "ptgMemAreaV" },    // 0x46 (70)
    { "OP_MEMERRV",     "ptgMemErrV" },     // 0x47 (71)
    { "OP_MEMNOMEMV",   "ptgMemNoMemV" },   // 0x48 (72)
    { "OP_MEMFUNCV",    "ptgMemFuncV" },    // 0x49 (73)
    { "OP_REFERRV",     "ptgRefErrV" },     // 0x4a (74)
    { "OP_AREAERRV",    "ptgAreaErrV" },    // 0x4b (75)
    { "OP_REFNV",       "ptgRefNV" },       // 0x4c (76)
    { "OP_AREANV",      "ptgAreaNV" },      // 0x4d (77)
    { "OP_MEMAREANV",   "ptgMemAreaNV" },   // 0x4e (78)
    { "OP_MEMNOMEMNV",  "ptgMemNoMemNV" },  // 0x4f (79)
    { "",               "" },               // 0x50 (80)
    { "",               "" },               // 0x51 (81)
    { "",               "" },               // 0x52 (82)
    { "",               "" },               // 0x53 (83)
    { "",               "" },               // 0x54 (84)
    { "",               "" },               // 0x55 (85)
    { "",               "" },               // 0x56 (86)
    { "",               "" },               // 0x57 (87)
    { "OP_FUNCCEV",     "ptgFuncCEV" },     // 0x58 (88)
    { "OP_NAMEXV",      "ptgNameXV" },      // 0x59 (89)
    { "OP_REF3DV",      "ptgRef3dV" },      // 0x5a (90)
    { "OP_AREA3DV",     "ptgArea3dV" },     // 0x5b (91)
    { "OP_REFERR3DV",   "ptgRefErr3dV" },   // 0x5c (92)
    { "OP_AREAERR3DV",  "ptgAreaErr3dV" },  // 0x5d (93)
    { "",               "" },               // 0x5e (94)
    { "",               "" },               // 0x5f (95)
    { "OP_ARRAYA",      "ptgArrayA" },      // 0x60 (96)
    { "OP_FUNCA",       "ptgFuncA" },       // 0x61 (97)
    { "OP_FUNCVARA",    "ptgFuncVarA" },    // 0x62 (98)
    { "OP_NAMEA",       "ptgNameA" },       // 0x63 (99)
    { "OP_REFA",        "ptgRefA" },        // 0x64 (100)
    { "OP_AREAA",       "ptgAreaA" },       // 0x65 (101)
    { "OP_MEMAREAA",    "ptgMemAreaA" },    // 0x66 (102)
    { "OP_MEMERRA",     "ptgMemErrA" },     // 0x67 (103)
    { "OP_MEMNOMEMA",   "ptgMemNoMemA" },   // 0x68 (104)
    { "OP_MEMFUNCA",    "ptgMemFuncA" },    // 0x69 (105)
    { "OP_REFERRA",     "ptgRefErrA" },     // 0x6a (106)
    { "OP_AREAERRA",    "ptgAreaErrA" },    // 0x6b (107)
    { "OP_REFNA",       "ptgRefNA" },       // 0x6c (108)
    { "OP_AREANA",      "ptgAreaNA" },      // 0x6d (109)
    { "OP_MEMAREANA",   "ptgMemAreaNA" },   // 0x6e (110)
    { "OP_MEMNOMEMNA",  "ptgMemNoMemNA" },  // 0x6f (111)
    { "",               "" },               // 0x70 (112)
    { "",               "" },               // 0x71 (113)
    { "",               "" },               // 0x72 (114)
    { "",               "" },               // 0x73 (115)
    { "",               "" },               // 0x74 (116)
    { "",               "" },               // 0x75 (117)
    { "",               "" },               // 0x76 (118)
    { "",               "" },               // 0x77 (119)
    { "OP_FUNCCEA",     "ptgFuncCEA" },     // 0x78 (120)
    { "OP_NAMEXA",      "ptgNameXA" },      // 0x79 (121)
    { "OP_REF3DA",      "ptgRef3dA" },      // 0x7a (122)
    { "OP_AREA3DA",     "ptgArea3dA" },     // 0x7b (123)
    { "OP_REFERR3DA",   "ptgRefErr3dA" },   // 0x7c (124)
    { "OP_AREAERR3DA",  "ptgAreaErr3dA" },  // 0x7d (125)
    { "",               "" },               // 0x7e (126)
    { "",               "" },               // 0x7f (127)
};

static const excelNames functionNames[368] = {
    { "FUNC_COUNT",             "COUNT" },              // 0x00 (0)
    { "FUNC_IF",                "IF" },                 // 0x01 (1)
    { "FUNC_ISNA",              "ISNA" },               // 0x02 (2)
    { "FUNC_ISERROR",           "ISERROR" },            // 0x03 (3)
    { "FUNC_SUM",               "SUM" },                // 0x04 (4)
    { "FUNC_AVERAGE",           "AVERAGE" },            // 0x05 (5)
    { "FUNC_MIN",               "MIN" },                // 0x06 (6)
    { "FUNC_MAX",               "MAX" },                // 0x07 (7)
    { "FUNC_ROW",               "ROW" },                // 0x08 (8)
    { "FUNC_COLUMN",            "COLUMN" },             // 0x09 (9)
    { "FUNC_NA",                "NA" },                 // 0x0a (10)
    { "FUNC_NPV",               "NPV" },                // 0x0b (11)
    { "FUNC_STDEV",             "STDEV" },              // 0x0c (12)
    { "FUNC_DOLLAR",            "DOLLAR" },             // 0x0d (13)
    { "FUNC_FIXED",             "FIXED" },              // 0x0e (14)
    { "FUNC_SIN",               "SIN" },                // 0x0f (15)
    { "FUNC_COS",               "COS" },                // 0x10 (16)
    { "FUNC_TAN",               "TAN" },                // 0x11 (17)
    { "FUNC_ATAN",              "ATAN" },               // 0x12 (18)
    { "FUNC_PI",                "PI" },                 // 0x13 (19)
    { "FUNC_SQRT",              "SQRT" },               // 0x14 (20)
    { "FUNC_EXP",               "EXP" },                // 0x15 (21)
    { "FUNC_LN",                "LN" },                 // 0x16 (22)
    { "FUNC_LOG10",             "LOG10" },              // 0x17 (23)
    { "FUNC_ABS",               "ABS" },                // 0x18 (24)
    { "FUNC_INT",               "INT" },                // 0x19 (25)
    { "FUNC_SIGN",              "SIGN" },               // 0x1a (26)
    { "FUNC_ROUND",             "ROUND" },              // 0x1b (27)
    { "FUNC_LOOKUP",            "LOOKUP" },             // 0x1c (28)
    { "FUNC_INDEX",             "INDEX" },              // 0x1d (29)
    { "FUNC_REPT",              "REPT" },               // 0x1e (30)
    { "FUNC_MID",               "MID" },                // 0x1f (31)
    { "FUNC_LEN",               "LEN" },                // 0x20 (32)
    { "FUNC_VALUE",             "VALUE" },              // 0x21 (33)
    { "FUNC_TRUE",              "TRUE" },               // 0x22 (34)
    { "FUNC_FALSE",             "FALSE" },              // 0x23 (35)
    { "FUNC_AND",               "AND" },                // 0x24 (36)
    { "FUNC_OR",                "OR" },                 // 0x25 (37)
    { "FUNC_NOT",               "NOT" },                // 0x26 (38)
    { "FUNC_MOD",               "MOD" },                // 0x27 (39)
    { "FUNC_DCOUNT",            "DCOUNT" },             // 0x28 (40)
    { "FUNC_DSUM",              "DSUM" },               // 0x29 (41)
    { "FUNC_DAVERAGE",          "DAVERAGE" },           // 0x2a (42)
    { "FUNC_DMIN",              "DMIN" },               // 0x2b (43)
    { "FUNC_DMAX",              "DMAX" },               // 0x2c (44)
    { "FUNC_DSTDEV",            "DSTDEV" },             // 0x2d (45)
    { "FUNC_VAR",               "VAR" },                // 0x2e (46)
    { "FUNC_DVAR",              "DVAR" },               // 0x2f (47)
    { "FUNC_TEXT",              "TEXT" },               // 0x30 (48)
    { "FUNC_LINEST",            "LINEST" },             // 0x31 (49)
    { "FUNC_TREND",             "TREND" },              // 0x32 (50)
    { "FUNC_LOGEST",            "LOGEST" },             // 0x33 (51)
    { "FUNC_GROWTH",            "GROWTH" },             // 0x34 (52)
    { "FUNC_GOTO",              "GOTO" },               // 0x35 (53)
    { "FUNC_HALT",              "HALT" },               // 0x36 (54)
    { "",                       "" },                   // 0x37 (55)
    { "FUNC_PV",                "PV" },                 // 0x38 (56)
    { "FUNC_FV",                "FV" },                 // 0x39 (57)
    { "FUNC_NPER",              "NPER" },               // 0x3a (58)
    { "FUNC_PMT",               "PMT" },                // 0x3b (59)
    { "FUNC_RATE",              "RATE" },               // 0x3c (60)
    { "FUNC_MIRR",              "MIRR" },               // 0x3d (61)
    { "FUNC_IRR",               "IRR" },                // 0x3e (62)
    { "FUNC_RAND",              "RAND" },               // 0x3f (63)
    { "FUNC_MATCH",             "MATCH" },              // 0x40 (64)
    { "FUNC_DATE",              "DATE" },               // 0x41 (65)
    { "FUNC_TIME",              "TIME" },               // 0x42 (66)
    { "FUNC_DAY",               "DAY" },                // 0x43 (67)
    { "FUNC_MONTH",             "MONTH" },              // 0x44 (68)
    { "FUNC_YEAR",              "YEAR" },               // 0x45 (69)
    { "FUNC_WEEKDAY",           "WEEKDAY" },            // 0x46 (70)
    { "FUNC_HOUR",              "HOUR" },               // 0x47 (71)
    { "FUNC_MINUTE",            "MINUTE" },             // 0x48 (72)
    { "FUNC_SECOND",            "SECOND" },             // 0x49 (73)
    { "FUNC_NOW",               "NOW" },                // 0x4a (74)
    { "FUNC_AREAS",             "AREAS" },              // 0x4b (75)
    { "FUNC_ROWS",              "ROWS" },               // 0x4c (76)
    { "FUNC_COLUMNS",           "COLUMNS" },            // 0x4d (77)
    { "FUNC_OFFSET",            "OFFSET" },             // 0x4e (78)
    { "FUNC_ABSREF",            "ABSREF" },             // 0x4f (79)
    { "FUNC_RELREF",            "RELREF" },             // 0x50 (80)
    { "FUNC_ARGUMENT",          "ARGUMENT" },           // 0x51 (81)
    { "FUNC_SEARCH",            "SEARCH" },             // 0x52 (82)
    { "FUNC_TRANSPOSE",         "TRANSPOSE" },          // 0x53 (83)
    { "FUNC_ERROR",             "ERROR" },              // 0x54 (84)
    { "FUNC_STEP",              "STEP" },               // 0x55 (85)
    { "FUNC_TYPE",              "TYPE" },               // 0x56 (86)
    { "FUNC_ECHO",              "ECHO" },               // 0x57 (87)
    { "FUNC_SETNAME",           "SETNAME" },            // 0x58 (88)
    { "FUNC_CALLER",            "CALLER" },             // 0x59 (89)
    { "FUNC_DEREF",             "DEREF" },              // 0x5a (90)
    { "FUNC_WINDOWS",           "WINDOWS" },            // 0x5b (91)
    { "FUNC_SERIES",            "SERIES" },             // 0x5c (92)
    { "FUNC_DOCUMENTS",         "DOCUMENTS" },          // 0x5d (93)
    { "FUNC_ACTIVECELL",        "ACTIVECELL" },         // 0x5e (94)
    { "FUNC_SELECTION",         "SELECTION" },          // 0x5f (95)
    { "FUNC_RESULT",            "RESULT" },             // 0x60 (96)
    { "FUNC_ATAN2",             "ATAN2" },              // 0x61 (97)
    { "FUNC_ASIN",              "ASIN" },               // 0x62 (98)
    { "FUNC_ACOS",              "ACOS" },               // 0x63 (99)
    { "FUNC_CHOOSE",            "CHOOSE" },             // 0x64 (100)
    { "FUNC_HLOOKUP",           "HLOOKUP" },            // 0x65 (101)
    { "FUNC_VLOOKUP",           "VLOOKUP" },            // 0x66 (102)
    { "FUNC_LINKS",             "LINKS" },              // 0x67 (103)
    { "FUNC_INPUT",             "INPUT" },              // 0x68 (104)
    { "FUNC_ISREF",             "ISREF" },              // 0x69 (105)
    { "FUNC_GETFORMULA",        "GETFORMULA" },         // 0x6a (106)
    { "FUNC_GETNAME",           "GETNAME" },            // 0x6b (107)
    { "FUNC_SETVALUE",          "SETVALUE" },           // 0x6c (108)
    { "FUNC_LOG",               "LOG" },                // 0x6d (109)
    { "FUNC_EXEC",              "EXEC" },               // 0x6e (110)
    { "FUNC_CHAR",              "CHAR" },               // 0x6f (111)
    { "FUNC_LOWER",             "LOWER" },              // 0x70 (112)
    { "FUNC_UPPER",             "UPPER" },              // 0x71 (113)
    { "FUNC_PROPER",            "PROPER" },             // 0x72 (114)
    { "FUNC_LEFT",              "LEFT" },               // 0x73 (115)
    { "FUNC_RIGHT",             "RIGHT" },              // 0x74 (116)
    { "FUNC_EXACT",             "EXACT" },              // 0x75 (117)
    { "FUNC_TRIM",              "TRIM" },               // 0x76 (118)
    { "FUNC_REPLACE",           "REPLACE" },            // 0x77 (119)
    { "FUNC_SUBSTITUTE",        "SUBSTITUTE" },         // 0x78 (120)
    { "FUNC_CODE",              "CODE" },               // 0x79 (121)
    { "FUNC_NAMES",             "NAMES" },              // 0x7a (122)
    { "FUNC_DIRECTORY",         "DIRECTORY" },          // 0x7b (123)
    { "FUNC_FIND",              "FIND" },               // 0x7c (124)
    { "FUNC_CELL",              "CELL" },               // 0x7d (125)
    { "FUNC_ISERR",             "ISERR" },              // 0x7e (126)
    { "FUNC_ISTEXT",            "ISTEXT" },             // 0x7f (127)
    { "FUNC_ISNUMBER",          "ISNUMBER" },           // 0x80 (128)
    { "FUNC_ISBLANK",           "ISBLANK" },            // 0x81 (129)
    { "FUNC_T",                 "T" },                  // 0x82 (130)
    { "FUNC_N",                 "N" },                  // 0x83 (131)
    { "FUNC_FOPEN",             "FOPEN" },              // 0x84 (132)
    { "FUNC_FCLOSE",            "FCLOSE" },             // 0x85 (133)
    { "FUNC_FSIZE",             "FSIZE" },              // 0x86 (134)
    { "FUNC_FREADLN",           "FREADLN" },            // 0x87 (135)
    { "FUNC_FREAD",             "FREAD" },              // 0x88 (136)
    { "FUNC_FWRITELN",          "FWRITELN" },           // 0x89 (137)
    { "FUNC_FWRITE",            "FWRITE" },             // 0x8a (138)
    { "FUNC_FPOS",              "FPOS" },               // 0x8b (139)
    { "FUNC_DATEVALUE",         "DATEVALUE" },          // 0x8c (140)
    { "FUNC_TIMEVALUE",         "TIMEVALUE" },          // 0x8d (141)
    { "FUNC_SLN",               "SLN" },                // 0x8e (142)
    { "FUNC_SYD",               "SYD" },                // 0x8f (143)
    { "FUNC_DDB",               "DDB" },                // 0x90 (144)
    { "FUNC_GETDEF",            "GETDEF" },             // 0x91 (145)
    { "FUNC_REFTEXT",           "REFTEXT" },            // 0x92 (146)
    { "FUNC_TEXTREF",           "TEXTREF" },            // 0x93 (147)
    { "FUNC_INDIRECT",          "INDIRECT" },           // 0x94 (148)
    { "FUNC_REGISTER",          "REGISTER" },           // 0x95 (149)
    { "FUNC_CALL",              "CALL" },               // 0x96 (150)
    { "FUNC_ADDBAR",            "ADDBAR" },             // 0x97 (151)
    { "FUNC_ADDMENU",           "ADDMENU" },            // 0x98 (152)
    { "FUNC_ADDCOMMAND",        "ADDCOMMAND" },         // 0x99 (153)
    { "FUNC_ENABLECOMMAND",     "ENABLECOMMAND" },      // 0x9a (154)
    { "FUNC_CHECKCOMMAND",      "CHECKCOMMAND" },       // 0x9b (155)
    { "FUNC_RENAMECOMMAND",     "RENAMECOMMAND" },      // 0x9c (156)
    { "FUNC_SHOWBAR",           "SHOWBAR" },            // 0x9d (157)
    { "FUNC_DELETEMENU",        "DELETEMENU" },         // 0x9e (158)
    { "FUNC_DELETECOMMAND",     "DELETECOMMAND" },      // 0x9f (159)
    { "FUNC_GETCHARTITEM",      "GETCHARTITEM" },       // 0xa0 (160)
    { "FUNC_DIALOGBOX",         "DIALOGBOX" },          // 0xa1 (161)
    { "FUNC_CLEAN",             "CLEAN" },              // 0xa2 (162)
    { "FUNC_MDETERM",           "MDETERM" },            // 0xa3 (163)
    { "FUNC_MINVERSE",          "MINVERSE" },           // 0xa4 (164)
    { "FUNC_MMULT",             "MMULT" },              // 0xa5 (165)
    { "FUNC_FILES",             "FILES" },              // 0xa6 (166)
    { "FUNC_IPMT",              "IPMT" },               // 0xa7 (167)
    { "FUNC_PPMT",              "PPMT" },               // 0xa8 (168)
    { "FUNC_COUNTA",            "COUNTA" },             // 0xa9 (169)
    { "FUNC_CANCELKEY",         "CANCELKEY" },          // 0xaa (170)
    { "",                       "" },                   // 0xab (171)
    { "",                       "" },                   // 0xac (172)
    { "",                       "" },                   // 0xad (173)
    { "",                       "" },                   // 0xae (174)
    { "FUNC_INITIATE",          "INITIATE" },           // 0xaf (175)
    { "FUNC_REQUEST",           "REQUEST" },            // 0xb0 (176)
    { "FUNC_POKE",              "POKE" },               // 0xb1 (177)
    { "FUNC_EXECUTE",           "EXECUTE" },            // 0xb2 (178)
    { "FUNC_TERMINATE",         "TERMINATE" },          // 0xb3 (179)
    { "FUNC_RESTART",           "RESTART" },            // 0xb4 (180)
    { "FUNC_HELP",              "HELP" },               // 0xb5 (181)
    { "FUNC_GETBAR",            "GETBAR" },             // 0xb6 (182)
    { "FUNC_PRODUCT",           "PRODUCT" },            // 0xb7 (183)
    { "FUNC_FACT",              "FACT" },               // 0xb8 (184)
    { "FUNC_GETCELL",           "GETCELL" },            // 0xb9 (185)
    { "FUNC_GETWORKSPACE",      "GETWORKSPACE" },       // 0xba (186)
    { "FUNC_GETWINDOW",         "GETWINDOW" },          // 0xbb (187)
    { "FUNC_GETDOCUMENT",       "GETDOCUMENT" },        // 0xbc (188)
    { "FUNC_DPRODUCT",          "DPRODUCT" },           // 0xbd (189)
    { "FUNC_ISNONTEXT",         "ISNONTEXT" },          // 0xbe (190)
    { "FUNC_GETNOTE",           "GETNOTE" },            // 0xbf (191)
    { "FUNC_NOTE",              "NOTE" },               // 0xc0 (192)
    { "FUNC_STDEVP",            "STDEVP" },             // 0xc1 (193)
    { "FUNC_VARP",              "VARP" },               // 0xc2 (194)
    { "FUNC_DSTDEVP",           "DSTDEVP" },            // 0xc3 (195)
    { "FUNC_DVARP",             "DVARP" },              // 0xc4 (196)
    { "FUNC_TRUNC",             "TRUNC" },              // 0xc5 (197)
    { "FUNC_ISLOGICAL",         "ISLOGICAL" },          // 0xc6 (198)
    { "FUNC_DCOUNTA",           "DCOUNTA" },            // 0xc7 (199)
    { "FUNC_DELETEBAR",         "DELETEBAR" },          // 0xc8 (200)
    { "FUNC_UNREGISTER",        "UNREGISTER" },         // 0xc9 (201)
    { "",                       "" },                   // 0xca (202)
    { "",                       "" },                   // 0xcb (203)
    { "FUNC_USDOLLAR",          "USDOLLAR" },           // 0xcc (204)
    { "FUNC_FINDB",             "FINDB" },              // 0xcd (205)
    { "FUNC_SEARCHB",           "SEARCHB" },            // 0xce (206)
    { "FUNC_REPLACEB",          "REPLACEB" },           // 0xcf (207)
    { "FUNC_LEFTB",             "LEFTB" },              // 0xd0 (208)
    { "FUNC_RIGHTB",            "RIGHTB" },             // 0xd1 (209)
    { "FUNC_MIDB",              "MIDB" },               // 0xd2 (210)
    { "FUNC_LENB",              "LENB" },               // 0xd3 (211)
    { "FUNC_ROUNDUP",           "ROUNDUP" },            // 0xd4 (212)
    { "FUNC_ROUNDDOWN",         "ROUNDDOWN" },          // 0xd5 (213)
    { "FUNC_ASC",               "ASC" },                // 0xd6 (214)
    { "FUNC_DBCS",              "DBCS" },               // 0xd7 (215)
    { "FUNC_RANK",              "RANK" },               // 0xd8 (216)
    { "",                       "" },                   // 0xd9 (217)
    { "",                       "" },                   // 0xda (218)
    { "FUNC_ADDRESS",           "ADDRESS" },            // 0xdb (219)
    { "FUNC_DAYS360",           "DAYS360" },            // 0xdc (220)
    { "FUNC_TODAY",             "TODAY" },              // 0xdd (221)
    { "FUNC_VDB",               "VDB" },                // 0xde (222)
    { "",                       "" },                   // 0xdf (223)
    { "",                       "" },                   // 0xe0 (224)
    { "",                       "" },                   // 0xe1 (225)
    { "",                       "" },                   // 0xe2 (226)
    { "FUNC_MEDIAN",            "MEDIAN" },             // 0xe3 (227)
    { "FUNC_SUMPRODUCT",        "SUMPRODUCT" },         // 0xe4 (228)
    { "FUNC_SINH",              "SINH" },               // 0xe5 (229)
    { "FUNC_COSH",              "COSH" },               // 0xe6 (230)
    { "FUNC_TANH",              "TANH" },               // 0xe7 (231)
    { "FUNC_ASINH",             "ASINH" },              // 0xe8 (232)
    { "FUNC_ACOSH",             "ACOSH" },              // 0xe9 (233)
    { "FUNC_ATANH",             "ATANH" },              // 0xea (234)
    { "FUNC_DGET",              "DGET" },               // 0xeb (235)
    { "FUNC_CREATEOBJECT",      "CREATEOBJECT" },       // 0xec (236)
    { "FUNC_VOLATILE",          "VOLATILE" },           // 0xed (237)
    { "FUNC_LASTERROR",         "LASTERROR" },          // 0xee (238)
    { "FUNC_CUSTOMUNDO",        "CUSTOMUNDO" },         // 0xef (239)
    { "FUNC_CUSTOMREPEAT",      "CUSTOMREPEAT" },       // 0xf0 (240)
    { "FUNC_FORMULACONVERT",    "FORMULACONVERT" },     // 0xf1 (241)
    { "FUNC_GETLINKINFO",       "GETLINKINFO" },        // 0xf2 (242)
    { "FUNC_TEXTBOX",           "TEXTBOX" },            // 0xf3 (243)
    { "FUNC_INFO",              "INFO" },               // 0xf4 (244)
    { "FUNC_GROUP",             "GROUP" },              // 0xf5 (245)
    { "FUNC_GETOBJECT",         "GETOBJECT" },          // 0xf6 (246)
    { "FUNC_DB",                "DB" },                 // 0xf7 (247)
    { "FUNC_PAUSE",             "PAUSE" },              // 0xf8 (248)
    { "",                       "" },                   // 0xf9 (249)
    { "",                       "" },                   // 0xfa (250)
    { "FUNC_RESUME",            "RESUME" },             // 0xfb (251)
    { "FUNC_FREQUENCY",         "FREQUENCY" },          // 0xfc (252)
    { "FUNC_ADDTOOLBAR",        "ADDTOOLBAR" },         // 0xfd (253)
    { "FUNC_DELETETOOLBAR",     "DELETETOOLBAR" },      // 0xfe (254)
    { "FUNC_UDF",               "" },                   // 0xff (255)
    { "FUNC_RESETTOOLBAR",      "RESETTOOLBAR" },       // 0x100 (256)
    { "FUNC_EVALUATE",          "EVALUATE" },           // 0x101 (257)
    { "FUNC_GETTOOLBAR",        "GETTOOLBAR" },         // 0x102 (258)
    { "FUNC_GETTOOL",           "GETTOOL" },            // 0x103 (259)
    { "FUNC_SPELLINGCHECK",     "SPELLINGCHECK" },      // 0x104 (260)
    { "FUNC_ERRORTYPE",         "ERRORTYPE" },          // 0x105 (261)
    { "FUNC_APPTITLE",          "APPTITLE" },           // 0x106 (262)
    { "FUNC_WINDOWTITLE",       "WINDOWTITLE" },        // 0x107 (263)
    { "FUNC_SAVETOOLBAR",       "SAVETOOLBAR" },        // 0x108 (264)
    { "FUNC_ENABLETOOL",        "ENABLETOOL" },         // 0x109 (265)
    { "FUNC_PRESSTOOL",         "PRESSTOOL" },          // 0x10a (266)
    { "FUNC_REGISTERID",        "REGISTERID" },         // 0x10b (267)
    { "FUNC_GETWORKBOOK",       "GETWORKBOOK" },        // 0x10c (268)
    { "FUNC_AVEDEV",            "AVEDEV" },             // 0x10d (269)
    { "FUNC_BETADIST",          "BETADIST" },           // 0x10e (270)
    { "FUNC_GAMMALN",           "GAMMALN" },            // 0x10f (271)
    { "FUNC_BETAINV",           "BETAINV" },            // 0x110 (272)
    { "FUNC_BINOMDIST",         "BINOMDIST" },          // 0x111 (273)
    { "FUNC_CHIDIST",           "CHIDIST" },            // 0x112 (274)
    { "FUNC_CHIINV",            "CHIINV" },             // 0x113 (275)
    { "FUNC_COMBIN",            "COMBIN" },             // 0x114 (276)
    { "FUNC_CONFIDENCE",        "CONFIDENCE" },         // 0x115 (277)
    { "FUNC_CRITBINOM",         "CRITBINOM" },          // 0x116 (278)
    { "FUNC_EVEN",              "EVEN" },               // 0x117 (279)
    { "FUNC_EXPONDIST",         "EXPONDIST" },          // 0x118 (280)
    { "FUNC_FDIST",             "FDIST" },              // 0x119 (281)
    { "FUNC_FINV",              "FINV" },               // 0x11a (282)
    { "FUNC_FISHER",            "FISHER" },             // 0x11b (283)
    { "FUNC_FISHERINV",         "FISHERINV" },          // 0x11c (284)
    { "FUNC_FLOOR",             "FLOOR" },              // 0x11d (285)
    { "FUNC_GAMMADIST",         "GAMMADIST" },          // 0x11e (286)
    { "FUNC_GAMMAINV",          "GAMMAINV" },           // 0x11f (287)
    { "FUNC_CEILING",           "CEILING" },            // 0x120 (288)
    { "FUNC_HYPGEOMDIST",       "HYPGEOMDIST" },        // 0x121 (289)
    { "FUNC_LOGNORMDIST",       "LOGNORMDIST" },        // 0x122 (290)
    { "FUNC_LOGINV",            "LOGINV" },             // 0x123 (291)
    { "FUNC_NEGBINOMDIST",      "NEGBINOMDIST" },       // 0x124 (292)
    { "FUNC_NORMDIST",          "NORMDIST" },           // 0x125 (293)
    { "FUNC_NORMSDIST",         "NORMSDIST" },          // 0x126 (294)
    { "FUNC_NORMINV",           "NORMINV" },            // 0x127 (295)
    { "FUNC_NORMSINV",          "NORMSINV" },           // 0x128 (296)
    { "FUNC_STANDARDIZE",       "STANDARDIZE" },        // 0x129 (297)
    { "FUNC_ODD",               "ODD" },                // 0x12a (298)
    { "FUNC_PERMUT",            "PERMUT" },             // 0x12b (299)
    { "FUNC_POISSON",           "POISSON" },            // 0x12c (300)
    { "FUNC_TDIST",             "TDIST" },              // 0x12d (301)
    { "FUNC_WEIBULL",           "WEIBULL" },            // 0x12e (302)
    { "FUNC_SUMXMY2",           "SUMXMY2" },            // 0x12f (303)
    { "FUNC_SUMX2MY2",          "SUMX2MY2" },           // 0x130 (304)
    { "FUNC_SUMX2PY2",          "SUMX2PY2" },           // 0x131 (305)
    { "FUNC_CHITEST",           "CHITEST" },            // 0x132 (306)
    { "FUNC_CORREL",            "CORREL" },             // 0x133 (307)
    { "FUNC_COVAR",             "COVAR" },              // 0x134 (308)
    { "FUNC_FORECAST",          "FORECAST" },           // 0x135 (309)
    { "FUNC_FTEST",             "FTEST" },              // 0x136 (310)
    { "FUNC_INTERCEPT",         "INTERCEPT" },          // 0x137 (311)
    { "FUNC_PEARSON",           "PEARSON" },            // 0x138 (312)
    { "FUNC_RSQ",               "RSQ" },                // 0x139 (313)
    { "FUNC_STEYX",             "STEYX" },              // 0x13a (314)
    { "FUNC_SLOPE",             "SLOPE" },              // 0x13b (315)
    { "FUNC_TTEST",             "TTEST" },              // 0x13c (316)
    { "FUNC_PROB",              "PROB" },               // 0x13d (317)
    { "FUNC_DEVSQ",             "DEVSQ" },              // 0x13e (318)
    { "FUNC_GEOMEAN",           "GEOMEAN" },            // 0x13f (319)
    { "FUNC_HARMEAN",           "HARMEAN" },            // 0x140 (320)
    { "FUNC_SUMSQ",             "SUMSQ" },              // 0x141 (321)
    { "FUNC_KURT",              "KURT" },               // 0x142 (322)
    { "FUNC_SKEW",              "SKEW" },               // 0x143 (323)
    { "FUNC_ZTEST",             "ZTEST" },              // 0x144 (324)
    { "FUNC_LARGE",             "LARGE" },              // 0x145 (325)
    { "FUNC_SMALL",             "SMALL" },              // 0x146 (326)
    { "FUNC_QUARTILE",          "QUARTILE" },           // 0x147 (327)
    { "FUNC_PERCENTILE",        "PERCENTILE" },         // 0x148 (328)
    { "FUNC_PERCENTRANK",       "PERCENTRANK" },        // 0x149 (329)
    { "FUNC_MODE",              "MODE" },               // 0x14a (330)
    { "FUNC_TRIMMEAN",          "TRIMMEAN" },           // 0x14b (331)
    { "FUNC_TINV",              "TINV" },               // 0x14c (332)
    { "",                       "" },                   // 0x14d (333)
    { "FUNC_MOVIECOMMAND",      "MOVIECOMMAND" },       // 0x14e (334)
    { "FUNC_GETMOVIE",          "GETMOVIE" },           // 0x14f (335)
    { "FUNC_CONCATENATE",       "CONCATENATE" },        // 0x150 (336)
    { "FUNC_POWER",             "POWER" },              // 0x151 (337)
    { "FUNC_PIVOTADDDATA",      "PIVOTADDDATA" },       // 0x152 (338)
    { "FUNC_GETPIVOTTABLE",     "GETPIVOTTABLE" },      // 0x153 (339)
    { "FUNC_GETPIVOTFIELD",     "GETPIVOTFIELD" },      // 0x154 (340)
    { "FUNC_GETPIVOTITEM",      "GETPIVOTITEM" },       // 0x155 (341)
    { "FUNC_RADIANS",           "RADIANS" },            // 0x156 (342)
    { "FUNC_DEGREES",           "DEGREES" },            // 0x157 (343)
    { "FUNC_SUBTOTAL",          "SUBTOTAL" },           // 0x158 (344)
    { "FUNC_SUMIF",             "SUMIF" },              // 0x159 (345)
    { "FUNC_COUNTIF",           "COUNTIF" },            // 0x15a (346)
    { "FUNC_COUNTBLANK",        "COUNTBLANK" },         // 0x15b (347)
    { "FUNC_SCENARIOGET",       "SCENARIOGET" },        // 0x15c (348)
    { "FUNC_OPTIONSLISTSGET",   "OPTIONSLISTSGET" },    // 0x15d (349)
    { "FUNC_ISPMT",             "ISPMT" },              // 0x15e (350)
    { "FUNC_DATEDIF",           "DATEDIF" },            // 0x15f (351)
    { "FUNC_DATESTRING",        "DATESTRING" },         // 0x160 (352)
    { "FUNC_NUMBERSTRING",      "NUMBERSTRING" },       // 0x161 (353)
    { "FUNC_ROMAN",             "ROMAN" },              // 0x162 (354)
    { "FUNC_OPENDIALOG",        "OPENDIALOG" },         // 0x163 (355)
    { "FUNC_SAVEDIALOG",        "SAVEDIALOG" },         // 0x164 (356)
    { "FUNC_VIEWGET",           "VIEWGET" },            // 0x165 (357)
    { "FUNC_GETPIVOTDATA",      "" },                   // 0x166 (358)
    { "FUNC_HYPERLINK",         "HYPERLINK" },          // 0x167 (359)
    { "FUNC_PHONETIC",          "PHONETIC" },           // 0x168 (360)
    { "FUNC_AVERAGEA",          "AVERAGEA" },           // 0x169 (361)
    { "FUNC_MAXA",              "MAXA" },               // 0x16a (362)
    { "FUNC_MINA",              "MINA" },               // 0x16b (363)
    { "FUNC_STDEVPA",           "STDEVPA" },            // 0x16c (364)
    { "FUNC_VARPA",             "VARPA" },              // 0x16d (365)
    { "FUNC_STDEVA",            "STDEVA" },             // 0x16e (366)
    { "FUNC_VARA",              "VARA" },               // 0x16f (367)
#if 0
    { "FUNC_BAHTTEXT",          "BAHTTEXT" },           // 0x170 (368)
    { "FUNC_THAIDAYOFWEEK",     "THAIDAYOFWEEK" },      // 0x171 (369)
    { "FUNC_THAIDIGIT",         "THAIDIGIT" },          // 0x172 (370)
    { "FUNC_THAIMONTHOFYEAR",   "THAIMONTHOFYEAR" },    // 0x173 (371)
    { "FUNC_THAINUMSOUND",      "THAINUMSOUND" },       // 0x174 (372)
    { "FUNC_THAINUMSTRING",     "THAINUMSTRING" },      // 0x175 (373)
    { "FUNC_THAISTRINGLENGTH",  "THAISTRINGLENGTH" },   // 0x176 (374)
    { "FUNC_ISTHAIDIGIT",       "ISTHAIDIGIT" },        // 0x177 (375)
    { "FUNC_ROUNDBAHTDOWN",     "ROUNDBAHTDOWN" },      // 0x178 (376)
    { "FUNC_ROUNDBAHTUP",       "ROUNDBAHTUP" },        // 0x179 (377)
    { "FUNC_THAIYEAR",          "THAIYEAR" },           // 0x17a (378)
    { "FUNC_RTD",               "RTD" },                // 0x17b (379)
    { "FUNC_CUBEVALUE",         "CUBEVALUE" },          // 0x17c (380)
    { "FUNC_CUBEMEMBER",        "CUBEMEMBER" },         // 0x17d (381)
    { "FUNC_CUBEMEMBERPROPERTY","CUBEMEMBERPROPERTY" }, // 0x17e (382)
    { "FUNC_CUBERANKEDMEMBER",  "CUBERANKEDMEMBER" },   // 0x17f (383)
    { "FUNC_HEX2BIN",           "HEX2BIN" },            // 0x180 (384)
    { "FUNC_HEX2DEC",           "HEX2DEC" },            // 0x181 (385)
    { "FUNC_HEX2OCT",           "HEX2OCT" },            // 0x182 (386)
    { "FUNC_DEC2BIN",           "DEC2BIN" },            // 0x183 (387)
    { "FUNC_DEC2HEX",           "DEC2HEX" },            // 0x184 (388)
    { "FUNC_DEC2OCT",           "DEC2OCT" },            // 0x185 (389)
    { "FUNC_OCT2BIN",           "OCT2BIN" },            // 0x186 (390)
    { "FUNC_OCT2HEX",           "OCT2HEX" },            // 0x187 (391)
    { "FUNC_OCT2DEC",           "OCT2DEC" },            // 0x188 (392)
    { "FUNC_BIN2DEC",           "BIN2DEC" },            // 0x189 (393)
    { "FUNC_BIN2OCT",           "BIN2OCT" },            // 0x18a (394)
    { "FUNC_BIN2HEX",           "BIN2HEX" },            // 0x18b (395)
    { "FUNC_IMSUB",             "IMSUB" },              // 0x18c (396)
    { "FUNC_IMDIV",             "IMDIV" },              // 0x18d (397)
    { "FUNC_IMPOWER",           "IMPOWER" },            // 0x18e (398)
    { "FUNC_IMABS",             "IMABS" },              // 0x18f (399)
    { "FUNC_IMSQRT",            "IMSQRT" },             // 0x190 (400)
    { "FUNC_IMLN",              "IMLN" },               // 0x191 (401)
    { "FUNC_IMLOG2",            "IMLOG2" },             // 0x192 (402)
    { "FUNC_IMLOG10",           "IMLOG10" },            // 0x193 (403)
    { "FUNC_IMSIN",             "IMSIN" },              // 0x194 (404)
    { "FUNC_IMCOS",             "IMCOS" },              // 0x195 (405)
    { "FUNC_IMEXP",             "IMEXP" },              // 0x196 (406)
    { "FUNC_IMARGUMENT",        "IMARGUMENT" },         // 0x197 (407)
    { "FUNC_IMCONJUGATE",       "IMCONJUGATE" },        // 0x198 (408)
    { "FUNC_IMAGINARY",         "IMAGINARY" },          // 0x199 (409)
    { "FUNC_IMREAL",            "IMREAL" },             // 0x19a (410)
    { "FUNC_COMPLEX",           "COMPLEX" },            // 0x19b (411)
    { "FUNC_IMSUM",             "IMSUM" },              // 0x19c (412)
    { "FUNC_IMPRODUCT",         "IMPRODUCT" },          // 0x19d (413)
    { "FUNC_SERIESSUM",         "SERIESSUM" },          // 0x19e (414)
    { "FUNC_FACTDOUBLE",        "FACTDOUBLE" },         // 0x19f (415)
    { "FUNC_SQRTPI",            "SQRTPI" },             // 0x1a0 (416)
    { "FUNC_QUOTIENT",          "QUOTIENT" },           // 0x1a1 (417)
    { "FUNC_DELTA",             "DELTA" },              // 0x1a2 (418)
    { "FUNC_GESTEP",            "GESTEP" },             // 0x1a3 (419)
    { "FUNC_ISEVEN",            "ISEVEN" },             // 0x1a4 (420)
    { "FUNC_ISODD",             "ISODD" },              // 0x1a5 (421)
    { "FUNC_MROUND",            "MROUND" },             // 0x1a6 (422)
    { "FUNC_ERF",               "ERF" },                // 0x1a7 (423)
    { "FUNC_ERFC",              "ERFC" },               // 0x1a8 (424)
    { "FUNC_BESSELJ",           "BESSELJ" },            // 0x1a9 (425)
    { "FUNC_BESSELK",           "BESSELK" },            // 0x1aa (426)
    { "FUNC_BESSELY",           "BESSELY" },            // 0x1ab (427)
    { "FUNC_BESSELI",           "BESSELI" },            // 0x1ac (428)
    { "FUNC_XIRR",              "XIRR" },               // 0x1ad (429)
    { "FUNC_XNPV",              "XNPV" },               // 0x1ae (430)
    { "FUNC_PRICEMAT",          "PRICEMAT" },           // 0x1af (431)
    { "FUNC_YIELDMAT",          "YIELDMAT" },           // 0x1b0 (432)
    { "FUNC_INTRATE",           "INTRATE" },            // 0x1b1 (433)
    { "FUNC_RECEIVED",          "RECEIVED" },           // 0x1b2 (434)
    { "FUNC_DISC",              "DISC" },               // 0x1b3 (435)
    { "FUNC_PRICEDISC",         "PRICEDISC" },          // 0x1b4 (436)
    { "FUNC_YIELDDISC",         "YIELDDISC" },          // 0x1b5 (437)
    { "FUNC_TBILLEQ",           "TBILLEQ" },            // 0x1b6 (438)
    { "FUNC_TBILLPRICE",        "TBILLPRICE" },         // 0x1b7 (439)
    { "FUNC_TBILLYIELD",        "TBILLYIELD" },         // 0x1b8 (440)
    { "FUNC_PRICE",             "PRICE" },              // 0x1b9 (441)
    { "FUNC_YIELD",             "YIELD" },              // 0x1ba (442)
    { "FUNC_DOLLARDE",          "DOLLARDE" },           // 0x1bb (443)
    { "FUNC_DOLLARFR",          "DOLLARFR" },           // 0x1bc (444)
    { "FUNC_NOMINAL",           "NOMINAL" },            // 0x1bd (445)
    { "FUNC_EFFECT",            "EFFECT" },             // 0x1be (446)
    { "FUNC_CUMPRINC",          "CUMPRINC" },           // 0x1bf (447)
    { "FUNC_CUMIPMT",           "CUMIPMT" },            // 0x1c0 (448)
    { "FUNC_EDATE",             "EDATE" },              // 0x1c1 (449)
    { "FUNC_EOMONTH",           "EOMONTH" },            // 0x1c2 (450)
    { "FUNC_YEARFRAC",          "YEARFRAC" },           // 0x1c3 (451)
    { "FUNC_COUPDAYBS",         "COUPDAYBS" },          // 0x1c4 (452)
    { "FUNC_COUPDAYS",          "COUPDAYS" },           // 0x1c5 (453)
    { "FUNC_COUPDAYSNC",        "COUPDAYSNC" },         // 0x1c6 (454)
    { "FUNC_COUPNCD",           "COUPNCD" },            // 0x1c7 (455)
    { "FUNC_COUPNUM",           "COUPNUM" },            // 0x1c8 (456)
    { "FUNC_COUPPCD",           "COUPPCD" },            // 0x1c9 (457)
    { "FUNC_DURATION",          "DURATION" },           // 0x1ca (458)
    { "FUNC_MDURATION",         "MDURATION" },          // 0x1cb (459)
    { "FUNC_ODDLPRICE",         "ODDLPRICE" },          // 0x1cc (460)
    { "FUNC_ODDLYIELD",         "ODDLYIELD" },          // 0x1cd (461)
    { "FUNC_ODDFPRICE",         "ODDFPRICE" },          // 0x1ce (462)
    { "FUNC_ODDFYIELD",         "ODDFYIELD" },          // 0x1cf (463)
    { "FUNC_RANDBETWEEN",       "RANDBETWEEN" },        // 0x1d0 (464)
    { "FUNC_WEEKNUM",           "WEEKNUM" },            // 0x1d1 (465)
    { "FUNC_AMORDEGRC",         "AMORDEGRC" },          // 0x1d2 (466)
    { "FUNC_AMORLINC",          "AMORLINC" },           // 0x1d3 (467)
    { "FUNC_CONVERT",           "CONVERT" },            // 0x1d4 (468)
    { "FUNC_ACCRINT",           "ACCRINT" },            // 0x1d5 (469)
    { "FUNC_ACCRINTM",          "ACCRINTM" },           // 0x1d6 (470)
    { "FUNC_WORKDAY",           "WORKDAY" },            // 0x1d7 (471)
    { "FUNC_NETWORKDAYS",       "NETWORKDAYS" },        // 0x1d8 (472)
    { "FUNC_GCD",               "GCD" },                // 0x1d9 (473)
    { "FUNC_MULTINOMIAL",       "MULTINOMIAL" },        // 0x1da (474)
    { "FUNC_LCM",               "LCM" },                // 0x1db (475)
    { "FUNC_FVSCHEDULE",        "FVSCHEDULE" },         // 0x1dc (476)
    { "FUNC_CUBEKPIMEMBER",     "CUBEKPIMEMBER" },      // 0x1dd (477)
    { "FUNC_CUBESET",           "CUBESET" },            // 0x1de (478)
    { "FUNC_CUBESETCOUNT",      "CUBESETCOUNT" },       // 0x1df (479)
    { "FUNC_IFERROR",           "IFERROR" },            // 0x1e0 (480)
    { "FUNC_COUNTIFS",          "COUNTIFS" },           // 0x1e1 (481)
    { "FUNC_SUMIFS",            "SUMIFS" },             // 0x1e2 (482)
    { "FUNC_AVERAGEIF",         "AVERAGEIF" },          // 0x1e3 (483)
    { "FUNC_AVERAGEIFS",        "AVERAGEIFS" },         // 0x1e4 (484)
    { "FUNC_AGGREGATE",         "AGGREGATE" },          // 0x1e5 (485)
    { "FUNC_BINOM_DIST",        "BINOM_DIST" },         // 0x1e6 (486)
    { "FUNC_BINOM_INV",         "BINOM_INV" },          // 0x1e7 (487)
    { "FUNC_CONFIDENCE_NORM",   "CONFIDENCE_NORM" },    // 0x1e8 (488)
    { "FUNC_CONFIDENCE_T",      "CONFIDENCE_T" },       // 0x1e9 (489)
    { "FUNC_CHISQ_TEST",        "CHISQ_TEST" },         // 0x1ea (490)
    { "FUNC_F_TEST",            "F_TEST" },             // 0x1eb (491)
    { "FUNC_COVARIANCE_P",      "COVARIANCE_P" },       // 0x1ec (492)
    { "FUNC_COVARIANCE_S",      "COVARIANCE_S" },       // 0x1ed (493)
    { "FUNC_EXPON_DIST",        "EXPON_DIST" },         // 0x1ee (494)
    { "FUNC_GAMMA_DIST",        "GAMMA_DIST" },         // 0x1ef (495)
    { "FUNC_GAMMA_INV",         "GAMMA_INV" },          // 0x1f0 (496)
    { "FUNC_MODE_MULT",         "MODE_MULT" },          // 0x1f1 (497)
    { "FUNC_MODE_SNGL",         "MODE_SNGL" },          // 0x1f2 (498)
    { "FUNC_NORM_DIST",         "NORM_DIST" },          // 0x1f3 (499)
    { "FUNC_NORM_INV",          "NORM_INV" },           // 0x1f4 (500)
    { "FUNC_PERCENTILE_EXC",    "PERCENTILE_EXC" },     // 0x1f5 (501)
    { "FUNC_PERCENTILE_INC",    "PERCENTILE_INC" },     // 0x1f6 (502)
    { "FUNC_PERCENTRANK_EXC",   "PERCENTRANK_EXC" },    // 0x1f7 (503)
    { "FUNC_PERCENTRANK_INC",   "PERCENTRANK_INC" },    // 0x1f8 (504)
    { "FUNC_POISSON_DIST",      "POISSON_DIST" },       // 0x1f9 (505)
    { "FUNC_QUARTILE_EXC",      "QUARTILE_EXC" },       // 0x1fa (506)
    { "FUNC_QUARTILE_INC",      "QUARTILE_INC" },       // 0x1fb (507)
    { "FUNC_RANK_AVG",          "RANK_AVG" },           // 0x1fc (508)
    { "FUNC_RANK_EQ",           "RANK_EQ" },            // 0x1fd (509)
    { "FUNC_STDEV_S",           "STDEV_S" },            // 0x1fe (510)
    { "FUNC_STDEV_P",           "STDEV_P" },            // 0x1ff (511)
    { "FUNC_T_DIST",            "T_DIST" },             // 0x200 (512)
    { "FUNC_T_DIST_2T",         "T_DIST_2T" },          // 0x201 (513)
    { "FUNC_T_DIST_RT",         "T_DIST_RT" },          // 0x202 (514)
    { "FUNC_T_INV",             "T_INV" },              // 0x203 (515)
    { "FUNC_T_INV_2T",          "T_INV_2T" },           // 0x204 (516)
    { "FUNC_VAR_S",             "VAR_S" },              // 0x205 (517)
    { "FUNC_VAR_P",             "VAR_P" },              // 0x206 (518)
    { "FUNC_WEIBULL_DIST",      "WEIBULL_DIST" },       // 0x207 (519)
    { "FUNC_NETWORKDAYS_INTL",  "NETWORKDAYS_INTL" },   // 0x208 (520)
    { "FUNC_WORKDAY_INTL",      "WORKDAY_INTL" },       // 0x209 (521)
    { "FUNC_ECMA_CEILING",      "ECMA_CEILING" },       // 0x20a (522)
    { "FUNC_ISO_CEILING",       "ISO_CEILING" },        // 0x20b (523)
    { "",                       "" },                   // 0x20c (524)
    { "FUNC_BETA_DIST",         "BETA_DIST" },          // 0x20d (525)
    { "FUNC_BETA_INV",          "BETA_INV" },           // 0x20e (526)
    { "FUNC_CHISQ_DIST",        "CHISQ_DIST" },         // 0x20f (527)
    { "FUNC_CHISQ_DIST_RT",     "CHISQ_DIST_RT" },      // 0x210 (528)
    { "FUNC_CHISQ_INV",         "CHISQ_INV" },          // 0x211 (529)
    { "FUNC_CHISQ_INV_RT",      "CHISQ_INV_RT" },       // 0x212 (530)
    { "FUNC_F_DIST",            "F_DIST" },             // 0x213 (531)
    { "FUNC_F_DIST_RT",         "F_DIST_RT" },          // 0x214 (532)
    { "FUNC_F_INV",             "F_INV" },              // 0x215 (533)
    { "FUNC_F_INV_RT",          "F_INV_RT" },           // 0x216 (534)
    { "FUNC_HYPGEOM_DIST",      "HYPGEOM_DIST" },       // 0x217 (535)
    { "FUNC_LOGNORM_DIST",      "LOGNORM_DIST" },       // 0x218 (536)
    { "FUNC_LOGNORM_INV",       "LOGNORM_INV" },        // 0x219 (537)
    { "FUNC_NEGBINOM_DIST",     "NEGBINOM_DIST" },      // 0x21a (538)
    { "FUNC_NORM_S_DIST",       "NORM_S_DIST" },        // 0x21b (539)
    { "FUNC_NORM_S_INV",        "NORM_S_INV" },         // 0x21c (540)
    { "FUNC_T_TEST",            "T_TEST" },             // 0x21d (541)
    { "FUNC_Z_TEST",            "Z_TEST" },             // 0x21e (542)
    { "FUNC_ERF_PRECISE",       "ERF_PRECISE" },        // 0x21f (543)
    { "FUNC_ERFC_PRECISE",      "ERFC_PRECISE" },       // 0x220 (544)
    { "FUNC_GAMMALN_PRECISE",   "GAMMALN_PRECISE" },    // 0x221 (545)
    { "FUNC_CEILING_PRECISE",   "CEILING_PRECISE" },    // 0x222 (546)
    { "FUNC_FLOOR_PRECISE",     "FLOOR_PRECISE" },      // 0x223 (547)
#endif
};


static int xls_is_bigendian()
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

static unsigned short xlsShortVal (short s)
{
    unsigned char c1, c2;
    
    if (xls_is_bigendian()) {
        c1 = s & 255;
        c2 = (s >> 8) & 255;
    
        return (c1 << 8) + c2;
    } else {
        return s & 0xFFFF;
    }
}

