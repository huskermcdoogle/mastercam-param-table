// The Summary page: built from a small dumped-style sheet, its cached order
// checked here, and written to summary_sample.xlsx for tools\check_summary.ps1
// to drive in real Excel (the live ranking, the batch, the insert cost).
//
// Usage: summary_test <out dir>
#include "../src/Summary.h"
#include "../src/Xlsx.h"

#include <cstdio>
#include <map>
#include <string>

namespace
	{
	int failed = 0;

	void Check (bool ok, const char *what)
		{
		std::printf ("  %s  %s\n", ok ? "ok   " : "FAIL ", what);
		if (!ok)
			++failed;
		}

	/// The summary row whose column B (or A) reads `t`, 1-based; 0 = none.
	size_t RowOf (const Xlsx::Sheet &s, const std::wstring &t)
		{
		for (size_t r = 0; r < s.summary.size (); ++r)
			for (size_t c = 0; c < 2 && c < s.summary[r].size (); ++c)
				if (s.summary[r][c].text == t)
					return r + 1;
		return 0;
		}

	const Xlsx::Sheet::FreeCell &At (const Xlsx::Sheet &s, size_t row, size_t col)
		{
		static const Xlsx::Sheet::FreeCell none;
		return row >= 1 && row - 1 < s.summary.size () && col < s.summary[row - 1].size () ? s.summary[row - 1][col] : none;
		}
	}

int main (int argc, char **argv)
	{
	const std::string dir = argc > 1 ? argv[1] : ".";
	Xlsx::Sheet s;
	//            A          B       C        D           E          F               G                  H             I          J
	s.rows = { { L"op_idn", L"type", L"tool", L"comment", L"changes", L"est_seconds", L"cut_seconds_est", L"flips_part", L"removed", L"air_pct" },
			   { L"2", L"ROUGH", L"1", L"Rough OD", L"0", L"300", L"240", L"2", L"10", L"25" },
			   { L"5", L"FINISH", L"3", L"Finish OD", L"0", L"120", L"100", L"0.5", L"1", L"" },
			   { L"7", L"GROOVE", L"3", L"Groove", L"0", L"300", L"200", L"1", L"2", L"50" },
			   { L"9", L"DRILL", L"", L"Spot", L"0", L"", L"", L"", L"", L"" } };
	s.group.assign (10, 0);
	s.groupNames = { L"Identity" };
	s.readOnly.assign (10, 1);
	s.text = { 0, 1, 0, 1, 0, 0, 0, 0, 0, 0 };
	s.formula.assign (4, std::vector<std::wstring> (10));
	s.frozenCols = 5;
	s.trackChanges = true;
	s.changesCol = 4;
	s.untracked.assign (10, 0);
	s.untracked[5] = s.untracked[6] = s.untracked[7] = 1;

	// The Tools page as the dump lays it out: D insert, E flips, F cut time; then
	// the inserts table (B insert, D edges, E flips, F inserts, G cost, H cost).
	{
	Xlsx::Sheet::ToolRow t1, t3;
	t1.number = L"1";
	t1.name = L"OD ROUGH";
	t1.usedBy = L"op 2";
	t3.number = L"3";
	t3.name = L"OD FINISH";
	t3.usedBy = L"op 5, 7";
	auto extras = [] (const wchar_t *row, const wchar_t *insert, const wchar_t *flips0, const wchar_t *cut0)
		{
		Xlsx::Sheet::FreeCell ins, flips, cut;
		ins.text = insert;
		ins.editable = true;
		flips.formula = std::wstring (L"SUMIF('Lathe params'!$C$3:$C$6,$A") + row + L",'Lathe params'!$H$3:$H$6)";
		flips.text = flips0;
		cut.formula = std::wstring (L"TEXT(SUMIF('Lathe params'!$C$3:$C$6,$A") + row
					  + L",'Lathe params'!$G$3:$G$6)/86400,\"[h]:mm:ss\")";
		cut.text = cut0;
		return std::vector<Xlsx::Sheet::FreeCell> { ins, flips, cut };
		};
	t1.extra = extras (L"2", L"CNMG 432", L"2", L"0:04:00");
	t3.extra = extras (L"3", L"VNMG 331", L"1.5", L"0:05:00");
	s.tools = { t1, t3 };
	s.toolExtraHeads = { L"Insert", L"Flips / part", L"Cut time / part" };

	auto head = [] (const wchar_t *t)
		{
		Xlsx::Sheet::FreeCell h;
		h.text = t;
		h.head = true;
		return h;
		};
	std::vector<Xlsx::Sheet::FreeCell> heads = { head (L"Inserts"), head (L"Insert"), head (L"Used by"),
												 head (L"Edges per insert"), head (L"Flips / part"), head (L"Inserts / part") };
	for (const std::wstring &h : Summary::CostHeads ())
		heads.push_back (head (h.c_str ()));
	s.toolsAfter.push_back (heads);
	// Tools rows 2-3, a blank row 4, the heading row 5, inserts on rows 6-7.
	const wchar_t *names[] = { L"CNMG 432", L"VNMG 331" };
	const wchar_t *edges[] = { L"4", L"2" };
	for (size_t i = 0; i < 2; ++i)
		{
		const std::wstring r = std::to_wstring (6 + i);
		Xlsx::Sheet::FreeCell blank, name, used, e, flips, inserts;
		name.text = names[i];
		name.editable = true;
		e.text = edges[i];
		e.editable = true;
		flips.formula = L"SUMIF($D$2:$D$3,$B" + r + L",$E$2:$E$3)";
		flips.text = i == 0 ? L"2" : L"1.5";
		inserts.formula = L"IF(N($D" + r + L")>0,ROUNDUP(ROUND($E" + r + L"/$D" + r + L",6),0),\"\")";
		inserts.text = L"1";
		std::vector<Xlsx::Sheet::FreeCell> row = { blank, name, used, e, flips, inserts };
		for (const Xlsx::Sheet::FreeCell &c : Summary::CostCells (6 + i))
			row.push_back (c);
		s.toolsAfter.push_back (row);
		}
	}

	Summary::Where w;
	w.title = L"Summary - SAMPLE.mcam";
	w.subtitle = L"Dumped 2026-10-09 12:00 - 4 operations";
	Summary::Add (s, w);

	Check (!s.summary.empty (), "a Summary page");
	Check (s.summaryWidths.size () == 9 && s.summaryWidths[8] == 0, "nine columns, the working column I hidden");

	// Longest: ops 2 and 7 tie at 300 s - sheet order breaks the tie (op 2 first).
	const size_t longest = RowOf (s, L"Longest operations (estimated time, now)");
	Check (longest > 0, "a longest-operations list");
	Check (At (s, longest + 2, 8).text == L"1" && At (s, longest + 3, 8).text == L"3" && At (s, longest + 4, 8).text == L"2",
		   "ranked 300 (op 2), 300 (op 7, tie after), 120");
	Check (At (s, longest + 2, 8).array && At (s, longest + 2, 8).formula.find (L"LARGE(") != std::wstring::npos,
		   "the rank is a live LARGE / MATCH array formula");
	Check (At (s, longest + 2, 1).text == L"Rough OD" && At (s, longest + 2, 2).text == L"2"
		   && At (s, longest + 2, 5).text == L"0:05:00", "first: Rough OD, op 2, 0:05:00");
	Check (At (s, longest + 5, 8).text.empty () && At (s, longest + 5, 1).text.empty (), "an op with no time is left off");
	Check (At (s, longest + 2, 2).look == Xlsx::Sheet::FreeCell::Link
		   && At (s, longest + 2, 2).formula.find (L"HYPERLINK(") != std::wstring::npos, "the op number links to its row");

	// Tools by flips: T1 (2) above T3 (1.5).
	const size_t tools = RowOf (s, L"Insert flips by tool (per part, now)");
	Check (tools > 0 && At (s, tools + 2, 1).text == L"OD ROUGH" && At (s, tools + 3, 1).text == L"OD FINISH",
		   "tools ranked by flips: OD ROUGH, then OD FINISH");

	// Air: op 7 (50% of 200 = 100 s) above op 2 (25% of 240 = 60 s).
	const size_t air = RowOf (s, L"Most air cutting (cut time with no stock in the way)");
	Check (air > 0 && At (s, air + 2, 1).text == L"Groove" && At (s, air + 2, 5).text == L"0:01:40"
		   && At (s, air + 3, 1).text == L"Rough OD", "air: Groove 0:01:40, then Rough OD");

	// Part totals.
	const size_t cycle = RowOf (s, L"Cycle time (estimate)");
	Check (cycle > 0 && At (s, cycle, 2).text == L"0:12:00" && At (s, cycle, 3).text == L"0:12:00",
		   "cycle time as dumped and now: 0:12:00");
	Check (At (s, cycle, 2).formula.find (L"Dumped!") != std::wstring::npos, "as dumped reads the Dumped sheet");
	Check (RowOf (s, L"Batch quantity (type it in)") > 0 && RowOf (s, L"Insert cost per batch (whole inserts)") > 0,
		   "a batch with its insert cost");
	Check (RowOf (s, L"How to use this workbook") > 0, "how to use it");

	const std::string out = dir + "/summary_sample.xlsx";
	if (!Xlsx::Write (out, s))
		{
		std::puts ("FAIL: write");
		return 1;
		}

	// The load still finds the main sheet, with the Summary in front of it.
	{
	std::vector<std::vector<std::wstring>> rows;
	std::vector<size_t> sheetRow;
	std::wstring why;
	const bool read = Xlsx::ReadSheet (out, rows, sheetRow, why);
	Check (read && rows.size () == 5 && rows[1][0] == L"2" && sheetRow[1] == 3, "the load reads 'Lathe params', not the Summary");
	std::map<std::string, std::string> parts;
	const std::string bytes = Xlsx::Build (s);
	Check (Xlsx::Unzip (bytes, parts, why), "unzips");
	const std::string &wb = parts["xl/workbook.xml"];
	Check (wb.find ("<sheet name=\"Summary\"") < wb.find ("<sheet name=\"Lathe params\""), "Summary is the first sheet");
	Check (wb.find ("localSheetId=\"1\"") != std::string::npos, "the filter names the main sheet by its new position");
	Check (parts["xl/worksheets/sheet4.xml"].find ("t=\"array\"") != std::string::npos, "array formulas written");
	Check (parts["xl/styles.xml"].find ("formatCode=\"0.0%\"") != std::string::npos, "a custom number format");
	}

	if (failed)
		{
		std::printf ("summary_test: %d FAILED\n", failed);
		return 1;
		}
	std::puts ("summary_test: all checks passed");
	return 0;
	}
