//
// Summary.h - the workbook's first page: what a programmer looks at to decide
// what to improve. No Mastercam SDK.
//
// Everything on it is a FORMULA over the other sheets, so it follows the edits:
// the part's totals as dumped and now, a batch (quantity typed in) with its
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
		};

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
