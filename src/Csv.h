//
// Csv.h - CSV text and cell values, with no Mastercam SDK in sight.
//
// This is the half of the tool where a round trip through Excel can go wrong
// silently: a number that comes back as "0.10000000000000001", a boolean Excel
// re-types, a comment with a comma in it that shifts every column after it.
// None of those crash - they write the wrong value into an operation. So the
// text handling is proven at a command prompt, before it is near a toolpath.
//
#pragma once

#include <string>
#include <vector>

namespace Csv
	{
	using Row = std::vector<std::wstring>;

	/// Rows to CSV text. CRLF line ends, every cell that needs it quoted,
	/// embedded quotes doubled. Excel's own dialect.
	std::wstring Serialize (const std::vector<Row> &rows);

	/// CSV text to rows. Tolerant of what Excel and Notepad each do to a file:
	/// a leading BOM, CRLF or bare LF, quoted cells holding commas, quotes and
	/// line breaks, and a missing final newline. Fully blank lines are skipped.
	std::vector<Row> Parse (const std::wstring &text);

	/// UTF-8 WITH a BOM, because Excel opens a BOM-less UTF-8 CSV as ANSI and
	/// mangles anything past ASCII - an operation comment with a degree sign
	/// would come back as three characters of garbage.
	std::string ToUtf8Bom (const std::wstring &text);
	std::wstring FromUtf8 (const std::string &bytes);

	// ---- Cell values -------------------------------------------------------

	/// The shortest text that parses back to EXACTLY this double. "0.1", never
	/// "0.10000000000000001", and never a rounded value that no longer equals
	/// what the operation held.
	std::wstring FormatDouble (double v);

	/// Spaces and tabs off both ends.
	std::wstring Trim (const std::wstring &s);

	/// A double as a PERSON reads it: the shortest decimal within half of
	/// SameDouble's tolerance, no exponent. Mastercam leaves noise in stored
	/// values (0.25 as 0.2500000000000021, 0.001 through a float as
	/// 0.0010000000474975); this drops it, and a dump loaded back untouched is
	/// still no change.
	std::wstring Tidy (double v);

	/// Strict: the whole cell must be a finite number. Surrounding spaces are
	/// forgiven; an empty cell, "NaN", "1.5abc" and "1,5" are not.
	bool ParseDouble (const std::wstring &text, double &out);

	/// Whole numbers only - "3" and "3.0" are accepted, "3.5" is not.
	bool ParseLong (const std::wstring &text, long long &out);

	/// 1/0, true/false, yes/no, on/off, case-insensitive.
	bool ParseBool (const std::wstring &text, bool &out);

	/// Two doubles equal to the precision a CSV round trip can be trusted to.
	bool SameDouble (double a, double b);
	}
