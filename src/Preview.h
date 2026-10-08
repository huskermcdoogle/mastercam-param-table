//
// Preview.h - the load's "here is exactly what will change" window.
//
// Every change a load would write, grouped by operation, old value struck
// through in red and new value in green; refused rows in amber with the
// reason; anything skipped or ignored listed too. Apply writes, Cancel does
// nothing - the same decision the plain Yes/No box used to ask, made where
// the whole change set can be read.
//
#pragma once

#include <string>
#include <vector>

namespace Preview
	{
	struct Line
		{
		enum Kind
			{
			Section,	//!< a heading across the list: text
			Op,			//!< one operation: text (op and type), detail (its comment)
			Change,		//!< one value: text (column), from, to
			Refused,	//!< one refused row: text (where), detail (why)
			Note		//!< plain information: text, detail
			};
		Kind kind = Note;
		std::wstring text;
		std::wstring detail;
		std::wstring from;
		std::wstring to;
		};

	/// Show the preview. `changes` is how many changes Apply would write; with
	/// none, the window only offers Close. True when the person chose Apply.
	bool Show (const std::wstring &title, const std::wstring &summary,
			   const std::vector<Line> &lines, int changes);
	}
