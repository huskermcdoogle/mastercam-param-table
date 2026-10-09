//
// Xlsx.h - write the dump as a formatted .xlsx. No Mastercam SDK.
//
// An xlsx is a zip of XML parts. The zip is written STORED (no compression),
// which needs only a CRC32, so there is no library. Text cells are inline
// strings, so there is no shared-strings part either.
//
// The sheet is a VIEW of the CSV: the same values, in the same columns, with
// a band of group names above the column names, frozen headers, widths,
// shading by operation kind, and grey for cells that cannot be edited. Nothing
// here changes a value - a number is written as a number only when its text
// is exactly that number, so the text Excel saves back out matches.
//
#pragma once

#include <filesystem>
#include <map>
#include <string>
#include <vector>

namespace Xlsx
	{
	struct Sheet
		{
		/// rows[0] is the header of column names; the rest are data rows.
		std::vector<std::vector<std::wstring>> rows;

		/// Per column: index into groupNames. Groups must be contiguous.
		std::vector<int> group;
		std::vector<std::wstring> groupNames;

		/// Per column: shown but never applied on a load - drawn grey.
		std::vector<char> readOnly;

		/// Per column: always a string, never a number (comment, type ...).
		std::vector<char> text;

		/// Per column: the cells are formatted as Text, so what is TYPED stays as
		/// typed - "9:00" is not turned into a time of day.
		std::vector<char> textFormat;

		/// Per DATA row, per column: the column does not apply to this row's
		/// kind of operation (blank in the CSV) - drawn grey.
		std::vector<std::vector<char>> notApplicable;

		/// Per DATA row, per column: a formula (no leading '=') that computes the
		/// cell, or empty. The cell's text in `rows` is the cached result - a
		/// number, or text for a formula that returns text. A formula cell is
		/// drawn grey - it is derived, not typed.
		std::vector<std::vector<std::wstring>> formula;

		/// Columns [0, frozenCols) stay in view when scrolling right.
		size_t frozenCols = 0;

		/// Which column holds the operation kind; rows are shaded by it.
		size_t kindCol = 1;

		/// Excel's filter dropdowns on the column-name row.
		bool autoFilter = true;

		/// WHAT A CELL ACCEPTS, checked by Excel as it is typed, and the tooltip
		/// shown when it is selected. One rule covers many cells.
		struct Validation
			{
			std::string cells;			//!< A1 ranges, space separated
			std::string type;			//!< "list", "decimal", "whole", "textLength", "custom", "" = tooltip only
			std::string op;				//!< "between", "greaterThanOrEqual", "lessThanOrEqual", "equal" (or "")
			std::wstring f1, f2;		//!< the limits / formula
			std::vector<std::wstring> choices;	//!< type "list": the choices (255 characters at most, joined)
			std::wstring title;			//!< tooltip title (Excel keeps 32 characters)
			std::wstring prompt;		//!< tooltip text (Excel keeps 255)
			std::wstring error;			//!< why an entry is refused (Excel keeps 225)
			};
		std::vector<Validation> validations;

		/// TRACK EDITS: a hidden sheet "Dumped" holds the values as written; every
		/// data cell that differs from it is highlighted. Rows are matched by
		/// column A (op_idn), so sorting or filtering does not confuse it.
		bool trackChanges = false;

		/// The column whose cells COUNT their row's edits (a formula), or -1.
		int changesCol = -1;

		/// Per column: left out of edit tracking - calculated columns whose value
		/// follows edits elsewhere (the live estimate) and is not an edit itself.
		std::vector<char> untracked;

		/// THE TOOLS PAGE: one row per tool, with its picture when there is one.
		/// Empty = no Tools sheet.
		/// A cell of the Tools page beyond the fixed columns: text, or a number,
		/// or a formula (no leading '=') with `text` its cached result.
		struct FreeCell
			{
			std::wstring text;
			std::wstring formula;
			bool head = false;			//!< drawn as a heading
			bool editable = false;		//!< drawn as a cell to type in (not grey)
			bool textFormat = false;	//!< formatted as Text: a typed 8:00 stays 8:00

			/// How a cell of a REPORT page looks. Auto is the grid's own: a heading,
			/// a cell to type in, or grey (worked out). The others are for a page
			/// that is read rather than typed in (the Summary): no grid colours.
			enum Look { Auto, Title, Section, Note, Plain, Link, Input };
			Look look = Auto;

			/// A number format: "" General, or "0%", "0.0%", "0.00", "currency" (the
			/// computer's own currency), "@" Text.
			std::wstring numFmt;

			/// An ARRAY formula (as if entered with Ctrl+Shift+Enter): array maths
			/// inside it works in every Excel, with or without dynamic arrays.
			bool array = false;

			/// Aligned right: a figure (or a time, which the sheet writes as text)
			/// lined up under its heading.
			bool right = false;
			};
		struct ToolRow
			{
			std::wstring number, name, usedBy;
			std::vector<FreeCell> extra;	//!< after "Used by", headed by toolExtraHeads
			std::string png;			//!< PNG bytes, or empty
			int width = 0, height = 0;	//!< the picture's pixels
			};
		std::vector<ToolRow> tools;
		std::vector<std::wstring> toolExtraHeads;

		/// Rows written under the tools, after one blank row (a summary table).
		std::vector<std::vector<FreeCell>> toolsAfter;

		/// THE SUMMARY PAGE: first in the workbook and the page it opens on - rows of
		/// cells from A1, no gridlines. Empty = no Summary sheet. Its formulas read
		/// the other sheets by name ('Lathe params', 'Tools', Dumped).
		std::vector<std::vector<FreeCell>> summary;
		/// Its column widths in characters, from A; 0 = a hidden working column.
		std::vector<double> summaryWidths;

		/// Main-sheet cells that link to a Tools row: (A1 cell, index into tools).
		std::vector<std::pair<std::string, size_t>> toolLinks;

		/// A MACRO WORKBOOK (.xlsm): the compiled VBA project (vbaProject.bin) and
		/// the ribbon tab's XML. Empty = a plain .xlsx. The compiled code expects
		/// the code names ThisWorkbook and Sheet1 (the main sheet); they are written.
		std::string vbaProject;
		std::string ribbonXml;

		/// Per group: gets an outline +/- (its first column stays visible).
		std::vector<char> outlineGroup;
		/// Per group: starts collapsed.
		std::vector<char> collapseGroup;
		};

	/// The bytes of the .xlsx file.
	std::string Build (const Sheet &s);

	/// Build and write to `file`. False when the file cannot be written.
	bool Write (const std::filesystem::path &file, const Sheet &s);

	// ---- Reading back what Excel saved ---------------------------------------
	//
	// Excel COMPRESSES every part when it saves, so reading needs DEFLATE - done
	// here (Inflate), not by a library. Cells come back as the text a CSV of the
	// same sheet would hold, so the load's rules apply unchanged.

	/// The rows of the sheet from the COLUMN-NAME row (the first row with a cell
	/// "op_idn") down; [0] is that row. Entirely empty rows are dropped, and
	/// `sheetRow` gives each returned row's row number in Excel. Formula cells
	/// read as the value Excel last calculated. False, with a reason, when the
	/// file is not a readable workbook or has no such row.
	bool ReadSheet (const std::filesystem::path &file,
					std::vector<std::vector<std::wstring>> &rows,
					std::vector<size_t> &sheetRow, std::wstring &why);

	/// The same, from the file's bytes.
	bool ReadSheetBytes (const std::string &bytes,
						 std::vector<std::vector<std::wstring>> &rows,
						 std::vector<size_t> &sheetRow, std::wstring &why);

	/// Another sheet of the workbook, by its exact name, read the same way -
	/// "Dumped", the values as written, is how the load knows what the edits
	/// moved the estimates FROM. False when there is no sheet of that name. An
	/// empty name is ReadSheet's own choice of sheet.
	bool ReadNamedSheet (const std::filesystem::path &file, const std::wstring &sheetName,
						 std::vector<std::vector<std::wstring>> &rows,
						 std::vector<size_t> &sheetRow, std::wstring &why);

	/// The same, from the file's bytes.
	bool ReadNamedSheetBytes (const std::string &bytes, const std::wstring &sheetName,
							  std::vector<std::vector<std::wstring>> &rows,
							  std::vector<size_t> &sheetRow, std::wstring &why);

	/// Raw DEFLATE (RFC 1951) data to bytes. False when the data is damaged.
	bool Inflate (const std::string &in, std::string &out);

	/// A zip's files by name, decompressed and CRC-checked.
	bool Unzip (const std::string &zip, std::map<std::string, std::string> &parts,
				std::wstring &why);

	/// Whether text is a plain decimal number this writer stores as a number.
	bool IsPlainNumber (const std::wstring &t);

	/// Zip helper, exposed for tests: a STORED zip of name/content pairs.
	std::string Zip (const std::vector<std::pair<std::string, std::string>> &parts);

	/// Column number to A1-style letters (0 -> "A", 26 -> "AA").
	std::string ColName (size_t col);
	}
