//
// Impact.h - what a sheet's edits do to the part: cycle time and insert flips,
// per operation and in total. No Mastercam SDK.
//
// The figures are the sheet's own. Excel keeps a formula's last result in the
// saved file, so the load reads the live estimate as the person last saw it:
//
//   time as dumped   cycle_time_raw   Mastercam's own cycle time, seconds
//   time now         est_seconds      the estimate with the edits, seconds
//   flips as dumped  flips_part       from the hidden "Dumped" sheet, which
//                                     holds every cell as it was written
//                    (flips)          when there is no Dumped sheet
//   flips now        flips_part       the live estimate of flips per part
//
// The two flip columns of the main sheet do not count the same thing (flips
// is the commented stops the toolpath makes; flips_part adds the ones the edge
// life implies), so "as dumped" is flips_part from the Dumped sheet whenever
// the workbook has it - otherwise an untouched sheet would show a change.
//
// An operation whose changes are all left out counts as dumped: its row may
// say otherwise, but nothing of it is written.
//
#pragma once

#include "Csv.h"

#include <cmath>
#include <map>
#include <set>
#include <string>
#include <vector>

namespace Impact
	{
	/// One operation's figures. NaN = the sheet does not say.
	struct Op
		{
		long op = 0;
		double was = std::nan ("");			//!< cycle time as dumped, seconds
		double now = std::nan ("");			//!< estimated with the edits, seconds
		double flipsWas = std::nan ("");	//!< insert flips per part as dumped
		double flipsNow = std::nan ("");	//!< with the edits
		};

	/// Every operation's figures from the main sheet (`sheet[0]` is the column
	/// names) and the Dumped sheet (may be empty). Rows without an op_idn are
	/// left out; a later row for the same op wins.
	std::map<long, Op> Read (const std::vector<Csv::Row> &sheet,
							 const std::vector<Csv::Row> &dumped);

	struct Total
		{
		double was = 0, now = 0;			//!< seconds
		double flipsWas = 0, flipsNow = 0;
		int timed = 0;						//!< operations with a time on both sides
		int flipped = 0;					//!< operations with flips on both sides
		};

	/// The whole part: every operation, those not in `applied` as dumped.
	Total Sum (const std::map<long, Op> &ops, const std::set<long> &applied);

	/// Whole seconds as h:mm:ss - "26:32:07", "0:04:12". Hours are not wrapped
	/// at 24: a cycle time is a length, not a time of day.
	std::wstring Hms (double seconds);

	/// A change in seconds with its sign: "-1:21:27", "+0:00:32".
	std::wstring Change (double seconds);

	/// -1 when `now` is better (less) than `was`, +1 when worse, 0 when the
	/// same to the second (or the hundredth of a flip, `flips` true).
	int Tone (double was, double now, bool flips = false);

	/// One operation's effect for its line in the preview: "0:04:12 -> 0:03:40
	/// (-0:00:32)  ·  flips 8 -> 6". Only what changed; "" when nothing did.
	std::wstring OpText (const Op &o);

	/// The whole part, for the top of the preview: "Cycle time 26:32:07 ->
	/// 25:10:40 (-1:21:27)  ·  flips 168 -> 152". "" when the sheet has no
	/// figures (a CSV, or a sheet dumped before the estimate existed).
	std::wstring TotalText (const Total &t);
	}
