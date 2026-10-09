//
// Inspect.h - lathe tool inspection: what the settings ask for, what the NCI
// actually did, and why each stop happened.
//
// THE NCI IS THE TRUTH. Mastercam writes a 1055 record (the inspection comment)
// wherever it stops the tool; Paths::Walk notes how much cutting - length and
// feed seconds - the tool had done by each one. The settings only explain them:
//
//   time      feed time since the last stop reached insp_time. Stopping only
//             BETWEEN cuts (insp_between_cuts) waits for the cut to end, so the
//             real interval runs over the setting; mid-cut stops come close to it
//             (a pass nearly done - within insp_min_cut - is finished first).
//   distance  cut length since the last stop reached insp_dist.
//   end       the stop at the end of the operation (insp_at_end).
//   cuts      after a number of cuts, the first cut, each depth / groove / section.
//
// The timer starts again with every operation: time a tool spends across
// several short operations never adds up to a stop.
//
// A FLIP is a stop whose comment says the insert is changed - "CHANGE/ROTATE
// INSERT" and the like. Other stops are inspections only and are not counted.
//
#pragma once

#include "Paths.h"

#include <string>
#include <vector>

namespace Inspect
	{
	/// Seconds as "m:ss" ("h:mm:ss" from an hour; tenths kept when there are any).
	std::wstring MinSec (double seconds);

	/// "9:00", "1:02:30", "540" (seconds) -> seconds. False when it is none of those.
	bool ParseMinSec (const std::wstring &text, double &seconds);

	/// Plan check for an m:ss cell.
	bool CheckMinSec (const std::wstring &text, std::wstring &why);

	/// The comment marks an insert change (flip / rotate / change).
	bool IsFlip (const std::wstring &comment);

	/// One operation's inspection settings, as the sheet has them.
	struct Settings
		{
		bool doStop = false, timeOn = false, distOn = false, cutsOn = false, firstCut = false;
		bool eachDepth = false, eachGroove = false, eachSection = false, atEnd = false;
		bool betweenCuts = true, commentOn = false;
		double time = 0, dist = 0, minCut = 0;
		long cuts = 0, sections = 0;
		std::wstring comment;
		};

	/// What the NCI did, explained.
	struct Result
		{
		int stops = 0;				//!< every inspection stop
		int flips = 0;				//!< the stops that change the insert
		int byTime = 0, byDist = 0, atEnd = 0, other = 0;	//!< the FLIPS, by cause
		double longest = 0;			//!< most feed seconds between flips (start and end count)
		double tail = 0;			//!< feed seconds after the last flip (all of it with none)
		std::wstring why;			//!< "time 8 + end 1"
		std::wstring mode;			//!< "between cuts" / "mid-cut ..."
		};

	Result Explain (const Settings &s, const Paths::Totals &t);

	/// When it stops, in words: "every 5 in of cut, at end" ("" when it never
	/// stops). `unit` is the part's length unit ("in" / "mm").
	std::wstring Criteria (const Settings &s, const std::wstring &unit);

	/// A log line: settings, stops and causes.
	std::wstring Describe (long opIdn, const Settings &s, const Result &r);
	}
