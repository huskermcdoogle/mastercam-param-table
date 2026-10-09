//
// MillMrr.h - a mill operation's theoretical metal removal rate, LIVE: the
// number at the dumped values and the Excel formula that follows the sheet.
// No Mastercam SDK - the dump hands it the row's cells by column name.
//
//   mrr = ae x ap x feed per minute          (in^3/min; cm^3/min metric)
//
//   ae, radial engagement   contour: the tool diameter (a full slot - the most
//                           a contour can take), or the multi-pass rough
//                           spacing when multi passes are on;
//                           dynamic mill: its stepover (Mastercam stores the
//                           distance - the percent in its dialog is worked
//                           from it).
//   ap, axial depth         the depth-cut rough step when depth cuts are on,
//                           else the whole depth: top of stock to depth.
//   feed per minute         the feed, times RPM when it is per rev; the feed
//                           rate override when the contour's is on.
//
#pragma once

#include <map>
#include <string>

namespace MillMrr
	{
	/// One cell of the row: what it holds and where it is ("M7").
	struct Cell
		{
		std::wstring text, ref;
		};

	/// The row's cells by column name. A column the row does not have is absent.
	using Cells = std::map<std::wstring, Cell>;

	/// Every column the calculation may read - the caller fills these that exist.
	const wchar_t *const kColumns[] = {
		L"speed", L"speed_mode", L"max_ss", L"feed", L"feed_mode", L"fr_override_on", L"fr_override",
		L"dcuts_on", L"dcut_rough", L"mcuts_on", L"mcut_rough_n", L"mcut_rough_amt", L"stepover",
		L"depth", L"depth_inc", L"top_stock", L"top_stock_inc" };

	struct Result
		{
		bool ok = false;			//!< false: nothing to say (no depth, no feed ...) - leave the cell blank
		double mrr = 0;				//!< at the dumped values, rounded to 3 places as the formula is
		std::wstring formula;		//!< without the leading '=', wrapped in IFERROR
		std::wstring basis;			//!< what it was worked from, in words
		};

	/// `type` is the sheet's kind ("CONTOUR", "DYNAMIC MILL"); toolDia the
	/// operation's tool diameter; mm a metric part.
	Result Compute (const std::wstring &type, bool mm, double toolDia, const Cells &cells);
	}
