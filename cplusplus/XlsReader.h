/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * This file is part of libxls -- A multiplatform, C library
 * for parsing Excel(TM) files.
 *
 * libxls is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * libxls is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with libxls.  If not, see <http://www.gnu.org/licenses/>.
 *
 * Copyright 2012 David Hoerl
 *
 */

#include <string>

#include <xls.h>

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
			colStr(""),
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
		xlsString			char2string(const uint8_t *ptr) const;
#if XLS_WIDE_STRINGS == 1
		bool				isAscii(const uint8_t *ptr) const;
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
}