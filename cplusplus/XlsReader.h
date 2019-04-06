/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * Copyright 2004 Komarov Valery
 * Copyright 2006 Christophe Leitienne
 * Copyright 2008-2017 David Hoerl
 * Copyright 2013 Bob Colbert
 * Copyright 2013-2018 Evan Miller
 *
 * This file is part of libxls -- A multiplatform, C/C++ library for parsing
 * Excel(TM) files.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *    1. Redistributions of source code must retain the above copyright notice,
 *    this list of conditions and the following disclaimer.
 *
 *    2. Redistributions in binary form must reproduce the above copyright
 *    notice, this list of conditions and the following disclaimer in the
 *    documentation and/or other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS ''AS
 * IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO,
 * THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
 * PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
 * EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
 * PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
 * OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
 * WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
 * OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 */

#include <string>
#include <limits>
#include <stdexcept>

// Inside namespace xls:

#include "../include/xls.h"


#ifndef UINT32_MAX
#define UINT32_MAX std::numeric_limits<uint32_t>::max()
#endif

#define XLS_WIDE_STRINGS 0

#if XLS_WIDE_STRINGS == 0
typedef std::string xlsString;
#else
#include <iconv.h>
typedef std::wstring xlsString;
#endif

namespace xls
{
	typedef enum { cellBlank=0, cellString, cellInteger, cellFloat, cellBool, cellError, cellUnknown } contentsType;

	struct cellContent {
		contentsType		type;
		char				colStr[3];	// String "A"..."Z", "AA"..."ZZ" (second char is either nil or a capital letter)
		uint32_t			col;		// 1 based
		uint16_t			row;		// 1 based
		xlsString			str;		// even for numbers these values are formatted as well as provided below
		union Val {
			long			l;
			double			d;
			bool			b;
			int32_t			e;
			
		Val(int x) { l = x; }
		~Val() { }
		} val;
		
		cellContent(void) :
			type(cellBlank),
			colStr(),
			col(0),
			row(0),
			val(0) { }
		~cellContent() {}
	};

	class WorkBook
	{
	public:
#if XLS_WIDE_STRINGS == 0
		// characterSet is the 8-bit encoding you want to convert 16-bit unicode strings to
		WorkBook(const std::string& fileName, int debug=0, const char *characterSet="UTF-8");	
#else
		// characterSet has to be UTF-8
		WorkBook(const std::string& fileName, int debug=0);	
#endif
		~WorkBook();

		std::string			GetLibraryVersion() const;
		uint32_t			GetSheetCount() const;

		// Sheets
		xlsString			GetSheetName(uint32_t sheetNum) const;
		bool				GetSheetVisible(uint32_t sheetNum) const;

		// Summary
		xlsString			GetSummaryAppName(void) const;
		xlsString			GetSummaryAuthor(void) const;
		xlsString			GetSummaryCategory(void) const;
		xlsString			GetSummaryComment(void) const;
		xlsString			GetSummaryCompany(void) const;
		xlsString			GetSummaryKeywords(void) const;
		xlsString			GetSummaryLastAuthor(void) const;
		xlsString			GetSummaryManager(void) const;
		xlsString			GetSummarySubject(void) const;
		xlsString			GetSummaryTitle(void) const;

		cellContent			GetCell(uint32_t workSheetIndex, uint16_t row, uint16_t col);			// uses 1 based indexing!
		cellContent			GetCell(uint32_t workSheetIndex, uint16_t row, const char *colStr);		// "A"...."Z" "AA"..."ZZ"

		void				InitIterator(uint32_t sheetNum = UINT32_MAX);							// call this first...
		cellContent			GetNextCell(void);														// ...then this continually til you get a blank cell

		void				ShowCell(const cellContent& content) const;

	private:
		WorkBook(const WorkBook& that);
		WorkBook& operator=(const WorkBook& right);

	private:
		void				OpenSheet(uint32_t sheetNum);
		void				FormatCell(xlsCell *cell, cellContent& content) const;
		xlsString			char2string(const char *ptr) const;
#if XLS_WIDE_STRINGS == 1
		bool				isAscii(const char *ptr) const;
#endif

	private:
		const char			*charSet;			// must be first ivar
		bool				isUTF8;				// unused in the wstring case
#if XLS_WIDE_STRINGS == 1
		iconv_t				iconvCD;
#endif
		uint32_t			numSheets;
		xlsWorkBook			*workBook;
		uint32_t			activeWorkSheetID;		// keep last one active
		xlsWorkSheet		*activeWorkSheet;	// keep last one active
		xlsSummaryInfo		*summary;
		
		bool				iterating;
		uint32_t			lastRowIndex;
		uint32_t			lastColIndex;
	};

	/*
	 *  This is the exception which can be thrown by the xls reader.
	 */
	class XlsException : public std::runtime_error
	{
	public:
		XlsException(const XlsException &ex) throw() 
			: std::runtime_error(ex)
		{;}
		explicit XlsException(const std::string &msg) throw() 
			: std::runtime_error("XlsReader: " + msg)
		{;}
		virtual ~XlsException() throw()
		{;}
	};
}
