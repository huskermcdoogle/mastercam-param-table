//
// Summary.h - the workbook's first page: what a programmer looks at to decide
// what to improve. No Mastercam SDK.
//
// Everything on it is a FORMULA over the other sheets, so it follows the edits:
// the part's totals as dumped and now, what the part really took on the
// machine (typed in) against the estimate, a batch (quantity typed in) with its
// insert cost, the longest operations, the tools that flip inserts most, and
// the operations that cut the most air.
//
// RANKING IS LIVE. Each list's order comes from LARGE / MATCH over the main
// sheet, with each row's position as the tie-breaker (key = value - row / 1e9,
// so equal values keep sheet order), in one hidden working column (I) holding
// the position of the n-th largest; the visible cells INDEX by it. Those cells
// are array formulas, so the maths works in every Excel - no dynamic arrays
// needed. The text written with them is the order at the dump.
//
#pragma once

#include "Xlsx.h"

#include <string>
#include <vector>

namespace Summary
	{
	/// What the Summary calls the part.
	struct Where
		{
		std::wstring title;			//!< the part, as a heading
		std::wstring subtitle;		//!< when it was dumped, how many operations
		std::wstring batchQty;		//!< remembered (the part's .ptconfig); "" = 1
		std::wstring manual;		//!< the user manual's front page (index.html), "" = none

		/// The newest measurement from the machine (the part's history), shown
		/// again in the "From the machine" cells; "" = none yet.
		std::wstring actualCycle;	//!< h:mm:ss
		std::wstring actualDate;	//!< an Excel date (days since 1899-12-30); "" = it was not typed
		std::wstring actualParts, actualInserts, actualNote;

		/// False: only some of the part's operations are on the sheet, so its
		/// estimate is not the part's - said beside the measurement.
		bool wholePart = true;
		};

	/// THE "FROM THE MACHINE" CELLS: what the part really took, typed in, against
	/// the estimate. The part's history (History.h) finds them again in a saved
	/// workbook by these labels - column B, the value in C - so they are named
	/// once, here.
	namespace Machine
		{
		const wchar_t *const kSection = L"From the machine";
		const wchar_t *const kCycle = L"Actual cycle time per part (m:ss or h:mm:ss)";
		const wchar_t *const kDate = L"Date it was measured";
		const wchar_t *const kParts = L"Parts run";
		const wchar_t *const kInserts = L"Inserts used for that run (total, if known)";
		const wchar_t *const kNote = L"Note about the run";
		const wchar_t *const kEstimate = L"Estimate now, per part";
		const wchar_t *const kAgainst = L"Machine against the estimate";
		const wchar_t *const kInsertsPer = L"Inserts per part: machine, estimate";
		}

	/// Fill s.summary and s.summaryWidths from the sheet as it will be written:
	/// s.rows (the main sheet, columns found by name - op_idn, type, tool,
	/// comment, est_seconds, cut_seconds_est, flips_part, removed, air_pct), and
	/// s.tools / s.toolExtraHeads / s.toolsAfter (the Tools page). Sections whose
	/// columns are missing are left out. The inserts table is found by its
	/// heading row in toolsAfter ("Inserts" in A, CostHeads from G): its rows run
	/// to the first empty row, columns B insert, D edges per insert, E flips /
	/// part, G cost per insert, H insert cost / part.
	void Add (Xlsx::Sheet &s, const Where &w);

	/// The inserts table's cost cells for one row: "Cost per insert" (typed, the
	/// computer's currency) and "Insert cost / part" (live: inserts used per part
	/// x cost - flips / edges, not rounded, so it is the average over a run).
	/// `row` is the Tools page row; D, E and G of it are edges, flips and cost.
	std::vector<Xlsx::Sheet::FreeCell> CostCells (size_t row);

	/// The headings of those two cells.
	std::vector<std::wstring> CostHeads ();

	/// Number of rows in each ranked list.
	const size_t kTop = 10;
	}
