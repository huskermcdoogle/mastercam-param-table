//
// Plan.h - what a reloaded CSV is allowed to change. No Mastercam SDK.
//
// A reload writes into toolpath parameters, and a wrong write there does not
// crash - it makes the next posted program cut differently. So the decisions
// are made here, where they can be tested, and the Mastercam-facing code is
// left with nothing to decide: it applies the plan or it does not.
//
// THE RULES, each of which exists because the alternative is a silent wrong
// value:
//
//  * A CELL LEFT EMPTY MEANS "LEAVE IT ALONE", never zero. Deleting a cell in
//    Excel is not a request to set a stock allowance to nothing.
//  * A ROW IS ADDRESSED BY op_idn AND CHECKED AGAINST ITS TYPE. Sorting or
//    re-pasting a sheet must not be able to hand one operation another's row,
//    and a rough row must not be applied to a finish operation.
//  * ONLY A CELL THAT DIFFERS IS A CHANGE. An untouched sheet writes nothing,
//    and a number Excel re-formatted without changing is not a change.
//  * A ROW WITH ANY UNPARSEABLE OR OUT-OF-RANGE CELL IS REFUSED WHOLE. Half of
//    a row applied is an operation in a state nobody chose.
//  * READ-ONLY COLUMNS (identity, and anything the tool will not change) are
//    reported if edited and never written.
//
#pragma once

#include "Csv.h"

#include <cmath>
#include <string>
#include <vector>

namespace Plan
	{
	enum class Type { Bool, Long, Double, Text };

	struct Col
		{
		std::wstring name;
		Type type = Type::Double;
		bool readOnly = false;

		/// Inclusive limits for Long and Double. NaN = unchecked on that side.
		double lo = std::nan ("");
		double hi = std::nan ("");

		/// Text only: a further check the host supplies (coolant names are the
		/// machine's own, known only inside Mastercam). Null = none.
		bool (*check) (const std::wstring &text, std::wstring &why) = nullptr;

		/// Text only: the ONLY values allowed, matched ignoring case ("per rev",
		/// "CW" ...). Empty = any text. The sheet offers them as a dropdown.
		std::vector<std::wstring> choices;
		};

	/// One kind of operation's editable columns, in file order.
	struct Schema
		{
		std::wstring type;			//!< "ROUGH" - what the `type` column says
		std::vector<Col> cols;

		/// Columns that exist on the SHEET but belong to a different kind of
		/// operation. One sheet carries every kind, so a rough row sits beside
		/// finish-only columns. They are left blank on a dump; a value typed
		/// into one is REFUSED rather than ignored - silently dropping an edit
		/// the person believes took effect is the failure this whole tool is
		/// built to avoid.
		std::vector<std::wstring> foreign;

		/// Index of a column by name, or -1.
		int Find (const std::wstring &name) const;
		};

	/// One operation as it stands NOW, as text aligned with Schema::cols.
	struct Current
		{
		long op = 0;
		std::wstring type;
		std::vector<std::wstring> values;
		};

	struct Change
		{
		long op = 0;
		size_t col = 0;				//!< index into Schema::cols
		std::wstring name;
		std::wstring from;
		std::wstring to;
		};

	struct Refusal
		{
		size_t line = 0;			//!< 1-based CSV line, header = 1
		long op = 0;				//!< 0 when the row has no usable op_idn
		std::wstring why;
		};

	struct Result
		{
		std::vector<Change> changes;
		std::vector<Refusal> refusals;

		int rows = 0;				//!< data rows examined
		int rowsChanged = 0;
		int rowsUnchanged = 0;
		int emptyCells = 0;			//!< left alone because they were empty

		/// Columns the file has that this schema does not - ignored, listed.
		std::vector<std::wstring> unknownColumns;
		};

	/// Compare a CSV against what the operations hold.
	///
	/// `csv[0]` is the header. It must hold `op_idn` and `type`; any schema
	/// column may be missing (that column is simply not compared) and any order
	/// is fine. A whole-file problem is one Refusal at line 1 with no changes.
	Result Make (const Schema &schema, const std::vector<Current> &current,
				 const std::vector<Csv::Row> &csv);

	/// Whether a cell's text is acceptable for a column. Empty text is the
	/// caller's business - this judges values.
	bool Valid (const Col &col, const std::wstring &text, std::wstring &why);

	/// A column the dump writes for the READER only - tool, comment, tool name,
	/// group name, units, the decoded coolant and so on. A load never compares
	/// or writes one, and does not list it as unknown.
	bool IsInfoColumn (const std::wstring &name);

	/// Two cell texts equal AS VALUES for this column type.
	bool SameValue (const Col &col, const std::wstring &a, const std::wstring &b);
	}
