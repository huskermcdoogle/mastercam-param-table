// The part's history (<part>.pthistory): its lines written and read back, times
// and dates, what a saved workbook says (the Summary's "From the machine" cells,
// the Change report's Why column) kept once each, insert cost per part from a
// workbook's sheets, the records a dump, load and regeneration write, and the
// History sheet - with history_sample.xlsx written for tools\check_summary.ps1
// to drive in real Excel.
//
// Usage: history_test <out dir>
//        history_test --harvest <workbook>   (prints the records a dump would add
//                                             from it - the check scripts use it)
#include "../src/History.h"
#include "../src/Summary.h"
#include "../src/Xlsx.h"

#include <cstdio>
#include <filesystem>
#include <fstream>
#include <iterator>
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

	std::string Narrow (const std::wstring &w)
		{
		return Csv::ToUtf8Bom (w).substr (3);
		}

	bool Near (double a, double b, double tol = 1e-6)
		{
		return std::fabs (a - b) <= tol;
		}

	using Cell = Xlsx::Sheet::FreeCell;

	Cell Head (const wchar_t *t)
		{
		Cell h;
		h.text = t;
		h.head = true;
		return h;
		}

	Cell Plain (const std::wstring &t)
		{
		Cell c;
		c.text = t;
		return c;
		}

	/// A dumped-style sheet: four ops, two tools, an inserts table (as the dump
	/// lays them out - see summary_test.cpp), costs 12.50 and 8.00.
	Xlsx::Sheet Sample ()
		{
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

		Xlsx::Sheet::ToolRow t1, t3;
		t1.number = L"1";
		t1.name = L"OD ROUGH";
		t1.usedBy = L"op 2";
		t3.number = L"3";
		t3.name = L"OD FINISH";
		t3.usedBy = L"op 5, 7";
		auto extras = [] (const wchar_t *row, const wchar_t *insert, const wchar_t *flips0)
			{
			Cell ins, flips;
			ins.text = insert;
			ins.editable = true;
			flips.formula = std::wstring (L"SUMIF('Lathe params'!$C$3:$C$6,$A") + row + L",'Lathe params'!$H$3:$H$6)";
			flips.text = flips0;
			return std::vector<Cell> { ins, flips };
			};
		t1.extra = extras (L"2", L"CNMG 432", L"2");
		t3.extra = extras (L"3", L"VNMG 331", L"1.5");
		s.tools = { t1, t3 };
		s.toolExtraHeads = { L"Insert", L"Flips / part" };

		std::vector<Cell> heads = { Head (L"Inserts"), Head (L"Insert"), Head (L"Used by"), Head (L"Edges per insert"),
									Head (L"Flips / part"), Head (L"Inserts / part") };
		for (const std::wstring &h : Summary::CostHeads ())
			heads.push_back (Head (h.c_str ()));
		s.toolsAfter.push_back (heads);
		// Tools rows 2-3, a blank row 4, the heading row 5, inserts on rows 6-7.
		const wchar_t *names[] = { L"CNMG 432", L"VNMG 331" };
		const wchar_t *edges[] = { L"4", L"2" };
		const wchar_t *flips[] = { L"2", L"1.5" };
		const wchar_t *costs[] = { L"12.5", L"8" };
		for (size_t i = 0; i < 2; ++i)
			{
			const std::wstring r = std::to_wstring (6 + i);
			Cell name = Plain (names[i]), e = Plain (edges[i]), f;
			f.formula = L"SUMIF($D$2:$D$3,$B" + r + L",$E$2:$E$3)";
			f.text = flips[i];
			std::vector<Cell> row = { Plain (L""), name, Plain (L""), e, f, Plain (L"1") };
			for (Cell c : Summary::CostCells (6 + i))
				{
				if (c.editable)
					c.text = costs[i];
				row.push_back (c);
				}
			s.toolsAfter.push_back (row);
			}
		return s;
		}

	std::string ReadAll (const std::filesystem::path &p)
		{
		std::ifstream in (p, std::ios::binary);
		return std::string ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		}

	/// The page row whose column B (or A) starts with `t`, 0-based; -1 = none.
	long RowOf (const Xlsx::Sheet::Page &pg, const std::wstring &t, size_t from = 0)
		{
		for (size_t r = from; r < pg.rows.size (); ++r)
			for (size_t c = 0; c < 2 && c < pg.rows[r].size (); ++c)
				if (pg.rows[r][c].text.compare (0, t.size (), t) == 0 && !t.empty ())
					return static_cast<long> (r);
		return -1;
		}

	std::wstring TextAt (const Xlsx::Sheet::Page &pg, long r, size_t c)
		{
		return r >= 0 && static_cast<size_t> (r) < pg.rows.size () && c < pg.rows[static_cast<size_t> (r)].size ()
				   ? pg.rows[static_cast<size_t> (r)][c].text : std::wstring ();
		}
	}

int main (int argc, char **argv)
	{
	if (argc > 2 && std::string (argv[1]) == "--harvest")
		{
		const std::filesystem::path wb (argv[2]);
		for (const History::Record &r : History::Measured (wb, {}, L"2026-10-13 07:05"))
			std::printf ("%s\n", Narrow (History::Line (r)).c_str ());
		for (const History::Record &r : History::Whys (wb, {}, L"2026-10-13 07:05"))
			std::printf ("%s\n", Narrow (History::Line (r)).c_str ());
		return 0;
		}
	const std::filesystem::path dir = argc > 1 ? argv[1] : ".";

	// ---- Lines: written and read back; a typed "|" and line break stay inside one field.
	{
	History::Record r;
	r.when = L"2026-10-10 11:02";
	r.kind = L"why";
	r.Set (L"op", L"12").Set (L"column", L"feed").Set (L"change", L"0.3 -> 0.25");
	r.Set (L"why", L"chatter | on the shoulder\r\nsee: the note");
	r.Set (L"empty", L"");
	const std::wstring line = History::Line (r);
	Check (line == L"2026-10-10 11:02 | why | op: 12 | column: feed | change: 0.3 -> 0.25 | why: chatter \u00A6 on the shoulder \u21B5 see: the note",
		   "a record as one line - the bar and the break kept out of the way, an empty field left off");
	History::Record back;
	Check (History::Parse (line, back) && back.when == r.when && back.kind == L"why" && back.fields.size () == 4
		   && back.Get (L"why") == L"chatter \u00A6 on the shoulder \u21B5 see: the note" && back.Get (L"change") == L"0.3 -> 0.25",
		   "read back: every field, a \": \" inside a value kept");
	Check (!History::Parse (L"# a comment", back) && !History::Parse (L"   ", back) && !History::Parse (L"2026-10-10", back),
		   "comments, blank lines and lines with no kind are not records");
	Check (History::Parse (L"2026-10-10|LOAD|Changes: 3", back) && back.kind == L"load" && back.Get (L"changes") == L"3",
		   "a line edited by hand (no spaces, capitals) still reads");
	}

	// ---- Times and dates.
	{
	Check (Near (History::Seconds (L"4:30"), 270) && Near (History::Seconds (L"1:02:30"), 3750)
		   && Near (History::Seconds (L"26:32:07"), 95527) && Near (History::Seconds (L"75:30"), 4530)
		   && Near (History::Seconds (L" 270 "), 270) && Near (History::Seconds (L"0.003125"), 270),
		   "times as typed: m:ss, h:mm:ss (past 24 hours), 75 minutes, seconds, Excel's fraction of a day");
	Check (std::isnan (History::Seconds (L"")) && std::isnan (History::Seconds (L"abc")) && std::isnan (History::Seconds (L"1:2:3:4"))
		   && std::isnan (History::Seconds (L"-5")), "not times: blank, words, four parts, below zero");
	Check (History::Hms (95527) == L"26:32:07" && History::Hms (252.4) == L"0:04:12" && History::Hms (std::nan ("")).empty (),
		   "h:mm:ss, hours not wrapped at 24");
	Check (History::Change (-4887) == L"-1:21:27" && History::Change (0.4) == L"+0:00:00", "a change with its sign");
	Check (History::DateOfSerial (46307) == L"2026-10-12" && History::DateOfSerial (45351) == L"2024-02-29"
		   && History::DateOfSerial (46307.75) == L"2026-10-12" && History::DateOfSerial (12).empty (),
		   "Excel dates: 46307 is 2026-10-12, a leap day, the time of day dropped, 1900 left alone");
	Check (Near (History::SerialOfDate (L"2026-10-12"), 46307) && Near (History::SerialOfDate (L"2024/02/29"), 45351)
		   && std::isnan (History::SerialOfDate (L"2026-02-30")) && std::isnan (History::SerialOfDate (L"12/10/2026")),
		   "dates to Excel's: yyyy-mm-dd or yyyy/mm/dd; no 30th of February; day/month order is not guessed");
	const auto sp = History::Span (L"26:32:07 -> 25:10:40", true);
	const auto one = History::Span (L"168", false);
	Check (Near (sp.first, 95527) && Near (sp.second, 90640) && Near (one.first, 168) && Near (one.second, 168),
		   "\"a -> b\" in two; one value is both");
	}

	// ---- The file: a header once, records added at the end, read back in order.
	{
	const std::filesystem::path file = dir / "history_test.pthistory";
	std::error_code ec;
	std::filesystem::remove (file, ec);
	History::Record a, b;
	a.when = L"2026-10-10 09:14";
	a.kind = L"dump";
	a.Set (L"ops", L"4");
	b.when = L"2026-10-10 11:02";
	b.kind = L"load";
	b.Set (L"changes", L"3");
	Check (History::Append (file, { a }) && History::Append (file, { b }), "appended twice");
	const std::string bytes = ReadAll (file);
	Check (bytes.rfind ("\xEF\xBB\xBF# Parameter Table Tool - ", 0) == 0 && bytes.find ("# Parameter Table Tool", 4) == std::string::npos
		   && bytes.find ("\xEF\xBB\xBF", 1) == std::string::npos && bytes.find ("\r\n2026-10-10 11:02 | load | changes: 3\r\n") != std::string::npos,
		   "a BOM and the header at the top, once; one record a line");
	// A line added by hand with no line break at the end: the next record is not glued to it.
	{
	std::ofstream out (file, std::ios::binary | std::ios::app);
	out << "2026-10-11 08:00 | note | text: by hand";
	}
	History::Book book;
	book.file = file;
	book.records = History::Read (file);
	History::Record c;
	c.when = L"2026-10-12 14:20";
	c.kind = L"regen";
	c.Set (L"ops", L"2");
	Check (book.Add ({ c }) && book.records.size () == 4, "a book adds to the file and to itself");
	const std::vector<History::Record> all = History::Read (file);
	Check (all.size () == 4 && all[0].kind == L"dump" && all[1].kind == L"load" && all[2].kind == L"note" && all[3].kind == L"regen",
		   "read back, oldest first, the hand-made line kept apart");
	Check (History::PathFor (L"C:\\parts\\shaft.mcam") == std::filesystem::path (L"C:\\parts\\shaft.pthistory"), "beside the part");
	}

	// ---- The "From the machine" cells: shown again, read back from the saved file, kept once.
	const std::filesystem::path wbFile = dir / "history_sample.xlsx";
	{
	Xlsx::Sheet s = Sample ();
	Summary::Where w;
	w.title = L"Summary - SAMPLE.mcam";
	w.subtitle = L"Dumped 2026-10-12 12:00 - 4 operations";
	w.actualCycle = L"0:13:10";
	w.actualDate = L"46307";
	w.actualParts = L"40";
	w.actualInserts = L"31";
	w.actualNote = L"first batch | after the change";
	Summary::Add (s, w);
	size_t cycleRow = 0, dateRow = 0;
	for (size_t r = 0; r < s.summary.size (); ++r)
		if (s.summary[r].size () > 2 && s.summary[r][1].text == Summary::Machine::kCycle)
			cycleRow = r + 1;
		else if (s.summary[r].size () > 2 && s.summary[r][1].text == Summary::Machine::kDate)
			dateRow = r + 1;
	Check (cycleRow > 0 && s.summary[cycleRow - 1][2].text == L"0:13:10" && s.summary[cycleRow - 1][2].textFormat
		   && s.summary[cycleRow - 1][2].editable, "the newest measurement shown again, the time cell formatted as Text");
	Check (s.summary[cycleRow - 1].size () > 8 && s.summary[cycleRow - 1][8].text == L"790"
		   && s.summary[cycleRow - 1][8].formula.find (L"CHOOSE(") != std::wstring::npos,
		   "its seconds in the hidden working column (790), worked out without TIMEVALUE");
	Check (dateRow == cycleRow + 1 && s.summary[dateRow - 1][2].numFmt == L"yyyy-mm-dd", "the date shown as yyyy-mm-dd");
	bool timeRule = false, dateRule = false;
	for (const Xlsx::Sheet::Validation &v : s.summaryValidations)
		{
		timeRule = timeRule || (v.cells == "C" + std::to_string (cycleRow) && v.type == "custom");
		dateRule = dateRule || (v.cells == "C" + std::to_string (dateRow) && v.type == "date");
		}
	Check (timeRule && dateRule, "the typed cells are checked as they are typed");

	// The History page of a few records, written into the same workbook.
	std::vector<History::Record> recs;
	History::Record d = History::DumpRecord (s, L"SAMPLE_params_1.xlsx", true, L"10", L"2026-10-12 12:00");
	recs.push_back (d);
	s.pages.push_back (History::Page (recs, L"SAMPLE.mcam", L"C:\\parts\\SAMPLE.pthistory", L"10"));
	Check (Xlsx::Write (wbFile, s), "written");

	const std::vector<History::Record> got = History::Measured (wbFile, {}, L"2026-10-13 07:05");
	Check (got.size () == 1 && got[0].kind == L"actual" && got[0].Get (L"measured") == L"2026-10-12" && got[0].Get (L"cycle") == L"0:13:10"
		   && got[0].Get (L"parts") == L"40" && got[0].Get (L"inserts") == L"31"
		   && got[0].Get (L"note") == L"first batch \u00A6 after the change" && got[0].Get (L"estimate") == L"0:12:00"
		   && !got[0].Has (L"date typed"),
		   "read back: date, time, parts, inserts, note, and the estimate beside it");
	Check (History::Measured (wbFile, got, L"2026-10-20 08:00").empty (), "the same measurement again is not a new one");

	// No date typed: the day it is read stands in, and the next dump leaves the cell blank again.
	Xlsx::Sheet s2 = Sample ();
	w.actualDate.clear ();
	Summary::Add (s2, w);
	const std::filesystem::path noDate = dir / "history_nodate.xlsx";
	Xlsx::Write (noDate, s2);
	const std::vector<History::Record> nd = History::Measured (noDate, got, L"2026-10-13 07:05");
	Check (nd.size () == 1 && nd[0].Get (L"measured") == L"2026-10-13" && nd[0].Get (L"date typed") == L"no",
		   "no date typed: the day it was read, and said so");
	Check (History::Measured (noDate, nd, L"2026-10-14 09:00").empty (), "read again another day: still the same measurement");
	const std::vector<History::Record> both = { nd[0], got[0] };
	const History::Record *newest = History::NewestActual (both);
	Check (newest != nullptr && newest->Get (L"measured") == L"2026-10-13", "the newest measurement is the latest date, not the last line");

	// Nothing typed: nothing kept.
	Xlsx::Sheet s3 = Sample ();
	Summary::Add (s3, Summary::Where ());
	const std::filesystem::path none = dir / "history_none.xlsx";
	Xlsx::Write (none, s3);
	Check (History::Measured (none, {}, L"2026-10-13 07:05").empty (), "no time typed: no measurement");
	}

	// ---- The Change report's Why column: one record per reason, kept once.
	{
	Xlsx::Sheet s = Sample ();
	Xlsx::Sheet::Page rep;
	rep.name = L"Change report";
	rep.rows.push_back ({ Plain (L"Change report - SAMPLE.mcam") });
	rep.rows.push_back ({ Plain (L"What"), Plain (L""), Plain (L""), Plain (L""), Plain (L"As dumped"), Plain (L"Now"), Plain (L"Change") });
	rep.rows.push_back ({ Plain (L"Op"), Plain (L"Tool"), Plain (L"Parameter"), Plain (L"column"), Plain (L"As dumped"), Plain (L"Now"),
						  Plain (L"Op time change"), Plain (L"Why (type the reason)"), Plain (L"key") });
	rep.rows.push_back ({ Plain (L"op 2"), Plain (L"T1"), Plain (L"ROUGH  -  Rough OD") });
	rep.rows.push_back ({ Plain (L""), Plain (L""), Plain (L"Feed"), Plain (L"feed"), Plain (L"0.012"), Plain (L"0.014"), Plain (L""),
						  Plain (L"tested on the floor"), Plain (L"2|feed") });
	rep.rows.push_back ({ Plain (L""), Plain (L""), Plain (L"Max RPM"), Plain (L"max_ss"), Plain (L"2000"), Plain (L"2500"), Plain (L""),
						  Plain (L""), Plain (L"2|max_ss") });
	s.pages.push_back (rep);
	const std::filesystem::path file = dir / "history_report.xlsx";
	Xlsx::Write (file, s);
	const std::vector<History::Record> w = History::Whys (file, {}, L"2026-10-10 11:02");
	Check (w.size () == 1 && w[0].kind == L"why" && w[0].Get (L"op") == L"2" && w[0].Get (L"column") == L"feed"
		   && w[0].Get (L"change") == L"0.012 -> 0.014" && w[0].Get (L"why") == L"tested on the floor",
		   "a reason typed: op, column, old -> new, why; a change with no reason left out");
	Check (History::Whys (file, w, L"2026-10-11 08:00").empty (), "the same reason again is not a new one");
	Check (History::Whys (wbFile, {}, L"2026-10-11 08:00").empty (), "no Change report: no reasons");
	}

	// ---- Insert cost per part, from the sheets the way the Tools page works it out.
	{
	const std::vector<Csv::Row> main = { { L"op_idn", L"tool", L"flips_part", L"xf_copies" },
										 { L"2", L"1", L"2", L"" },
										 { L"5", L"3", L"1", L"" },
										 { L"7", L"3", L"0.5", L"1" },
										 { L"9", L"4", L"0", L"" } };
	const std::vector<Csv::Row> dumped = { { L"op_idn", L"tool", L"flips_part" },
										   { L"2", L"1", L"3" },
										   { L"5", L"3", L"1" },
										   { L"7", L"3", L"0.5" },
										   { L"9", L"4", L"0" } };
	Xlsx::Grid tools;
	tools[1] = { { 0, L"Tool" }, { 1, L"Name" }, { 2, L"Used by" }, { 3, L"Insert" }, { 4, L"Flips / part" } };
	tools[2] = { { 0, L"1" }, { 3, L"CNMG 432" } };
	tools[3] = { { 0, L"3" }, { 3, L"VNMG 331" } };
	tools[4] = { { 0, L"4" }, { 3, L"RCMT 10" } };
	tools[6] = { { 0, L"Inserts" }, { 1, L"Insert" }, { 3, L"Edges per insert" }, { 4, L"Flips / part" }, { 6, L"Cost per insert" },
				 { 8, L"Parts per edge" } };
	tools[7] = { { 1, L"CNMG 432" }, { 3, L"4" }, { 6, L"12.5" } };
	tools[8] = { { 1, L"VNMG 331" }, { 3, L"2" }, { 6, L"8" } };
	tools[9] = { { 1, L"RCMT 10" }, { 3, L"8" }, { 6, L"20" }, { 8, L"50" } };
	// Now: CNMG 2/4 x 12.50 = 6.25; VNMG (1 + 0.5 x 2 copies) / 2 x 8 = 8; RCMT 1/50 / 8 x 20 = 0.05.
	const double now = History::InsertCost (main, dumped, tools, { 2, 5, 7, 9 });
	Check (Near (now, 14.3), "now: 6.25 + 8 + 0.05 (a transform's copy counted, parts per edge for the round)");
	// As dumped: CNMG 3 flips -> 9.375.
	Check (Near (History::InsertCost (main, dumped, tools, {}), 17.425), "as dumped: the Dumped sheet's flips");
	Check (Near (History::InsertCost (main, dumped, tools, { 5 }), 17.425), "an op not written counts as dumped");
	Check (Near (History::InsertCost (main, {}, tools, {}), 14.3), "no Dumped sheet: the main sheet's flips");
	Xlsx::Grid noCost = tools;
	noCost[7].erase (6);
	noCost[8].erase (6);
	noCost[9].erase (6);
	Check (std::isnan (History::InsertCost (main, dumped, noCost, {})), "no cost typed: none");

	// The same from a sheet about to be written.
	const Xlsx::Sheet s = Sample ();
	const Xlsx::Grid g = History::ToolsGrid (s);
	Check (g.count (1) && g.at (1).at (3) == L"Insert" && g.at (2).at (0) == L"1" && g.at (2).at (3) == L"CNMG 432"
		   && g.at (5).at (0) == L"Inserts" && g.at (6).at (1) == L"CNMG 432" && g.at (6).at (6) == L"12.5",
		   "the Tools page of a sheet to write, laid out as it will be read back");
	const History::Record d = History::DumpRecord (s, L"SAMPLE_params_1.xlsx", true, L"10", L"2026-10-12 12:00");
	// 2 / 4 x 12.50 + (0.5 + 1) / 2 x 8 = 6.25 + 6 = 12.25
	Check (d.kind == L"dump" && d.Get (L"ops") == L"4" && d.Get (L"whole part") == L"yes" && d.Get (L"cycle") == L"0:12:00"
		   && d.Get (L"flips") == L"3.5" && d.Get (L"insert cost") == L"12.25" && d.Get (L"batch") == L"10"
		   && d.Get (L"file") == L"SAMPLE_params_1.xlsx", "a dump's record: ops, cycle 0:12:00, flips 3.5, insert cost 12.25, batch");
	}

	// ---- A load's record, and a regeneration's.
	{
	History::LoadFigures f;
	f.files = L"SAMPLE_params_1.xlsx";
	f.changes = 3;
	f.ops = 2;
	f.cycleWas = 720;
	f.cycleNow = 660;
	f.flipsWas = 3.5;
	f.flipsNow = 3;
	f.costWas = 12.25;
	f.costNow = 11.5;
	f.batch = L"10";
	f.estimates = { { 2, 270 }, { 7, std::nan ("") } };
	const History::Record r = History::LoadRecord (f, L"2026-10-12 13:00");
	Check (History::Line (r) == L"2026-10-12 13:00 | load | file: SAMPLE_params_1.xlsx | changes: 3 | ops: 2 | cycle: 0:12:00 -> 0:11:00 "
								L"| flips: 3.5 -> 3 | insert cost: 12.25 -> 11.5 | batch: 10 | estimates: 2 0:04:30, 7 -",
		   "a load's line: before -> after, each op's estimate after");
	const std::map<long, double> e = History::Estimates (r);
	Check (e.size () == 2 && Near (e.at (2), 270) && std::isnan (e.at (7)), "its estimates read back");

	std::vector<History::Regenerated> ops (2);
	ops[0].op = 2;
	ops[0].estimate = 270;
	ops[0].mastercam = 276;
	ops[1].op = 7;
	ops[1].mastercam = 290;
	const std::vector<History::Record> g = History::RegenRecords (ops, 660, L"2026-10-12 14:20");
	Check (g.size () == 3 && g[0].Get (L"ops") == L"2" && g[0].Get (L"estimate") == L"0:04:30" && g[0].Get (L"mastercam") == L"0:04:36"
		   && g[0].Get (L"compared") == L"1" && g[0].Get (L"part estimate") == L"0:11:00" && g[0].Get (L"part mastercam") == L"0:11:06",
		   "a regeneration: the ops together (those with an estimate), the part with them at Mastercam's figures");
	Check (g[1].Get (L"op") == L"2" && g[1].Get (L"mastercam") == L"0:04:36" && g[2].Get (L"op") == L"7" && !g[2].Has (L"estimate"),
		   "then one line per op");
	}

	// ---- The History sheet: newest first, what each load saved, a running total, an undo.
	{
	auto rec = [] (const std::wstring &line)
		{
		History::Record r;
		History::Parse (line, r);
		return r;
		};
	const std::vector<History::Record> recs = {
		rec (L"2026-10-10 09:14 | dump | file: a.xlsx | ops: 42 | whole part: yes | cycle: 26:32:07 | flips: 168 | insert cost: 12.25 | batch: 40"),
		rec (L"2026-10-10 11:02 | load | file: a.xlsx | changes: 14 | ops: 6 | cycle: 26:32:07 -> 25:10:40 | flips: 168 -> 152 | insert cost: 12.25 -> 11.1 | batch: 40 | estimates: 12 0:04:12"),
		rec (L"2026-10-10 11:02 | why | op: 12 | column: feed | change: 0.3 -> 0.25 | why: chatter on the shoulder"),
		rec (L"2026-10-11 10:00 | load | file: b.xlsx | changes: 2 | ops: 1 | cycle: 25:10:40 -> 25:00:40 | batch: 40"),
		rec (L"2026-10-11 10:30 | undo | load: 2026-10-11 10:00:12 | restored: 2 | complete: yes"),
		rec (L"2026-10-12 14:20 | regen | ops: 1 | estimate: 0:04:12 | mastercam: 0:04:15 | part estimate: 25:10:40 | part mastercam: 25:10:43"),
		rec (L"2026-10-12 14:20 | regen | op: 12 | estimate: 0:04:12 | mastercam: 0:04:15"),
		rec (L"2026-10-13 07:05 | actual | measured: 2026-10-12 | cycle: 25:40:00 | parts: 40 | inserts: 31 | estimate: 25:10:40 | note: first batch") };
	const Xlsx::Sheet::Page pg = History::Page (recs, L"shaft.mcam", L"C:\\parts\\shaft.pthistory", L"40");
	Check (pg.name == L"History" && TextAt (pg, 0, 0) == L"History - shaft.mcam"
		   && TextAt (pg, 1, 0).find (L"shaft.pthistory") != std::wstring::npos
		   && TextAt (pg, 1, 0).find (L"C:\\parts") == std::wstring::npos, "titled for the part; names its file, not the folder");
	const long saved = RowOf (pg, L"2 loads (1 undone). Time saved per part: 1:21:27");
	Check (saved > 0 && TextAt (pg, saved, 0).find (L"5.1% of the first dump's 26:32:07") != std::wstring::npos,
		   "so far: the undone load taken back - 1:21:27 a part, 5.1% of the first dump");
	Check (RowOf (pg, L"Insert cost per part: 1.15 less - 46.00 a batch of 40.") > 0 && RowOf (pg, L"Insert flips per part: 16 fewer.") > 0,
		   "so far: inserts and their cost");
	Check (RowOf (pg, L"Last measured on the machine (2026-10-12): 25:40:00 a part - 0:29:20 slower than the estimate of 25:10:40 (+1.9%).") > 0,
		   "the last measurement against the estimate");
	Check (RowOf (pg, L"Last regeneration (2026-10-12 14:20): Mastercam was 0:00:03 slower than the estimate of 0:04:12 (+1.2%)") > 0,
		   "the last regeneration against the estimate");
	const long head = RowOf (pg, L"When");
	Check (head > 0 && pg.titleRow == static_cast<size_t> (head) + 1 && TextAt (pg, head, 3) == L"Saved / part",
		   "a heading row, printed on every page");
	// Newest first: the measurement, the regeneration (its op under it), the undo, the
	// second load (undone), the first load (its reason under it), the dump.
	const long actual = RowOf (pg, L"Measured on the machine on 2026-10-12: 40 parts");
	const long regen = RowOf (pg, L"Regenerated 1 operation");
	const long undo = RowOf (pg, L"Undid the last load: 2 values put back");
	const long load1 = RowOf (pg, L"Loaded 14 changes on 6 operations");
	const long load2 = RowOf (pg, L"Loaded 2 changes on 1 operation");
	const long dump = RowOf (pg, L"Dumped 42 operations (the whole part)");
	Check (actual > head && actual < regen && regen < undo && undo < load2 && load2 < load1 && load1 < dump, "newest first");
	Check (TextAt (pg, actual, 2) == L"25:40:00" && RowOf (pg, L"    31 inserts used - 0.78 a part", actual) == actual + 1
		   && RowOf (pg, L"    the machine was 0:29:20 slower than the estimate of 25:10:40 (+1.9%)", actual) == actual + 2,
		   "a measurement: the time, the inserts a part, against the estimate");
	Check (RowOf (pg, L"    op 12: Mastercam was 0:00:03 slower than the estimate of 0:04:12 (+1.2%)", regen) > regen
		   && TextAt (pg, regen, 2) == L"25:10:43 (Mastercam)", "a regeneration: the part at Mastercam's figures, each op under it");
	Check (TextAt (pg, load1, 2) == L"26:32:07 -> 25:10:40" && TextAt (pg, load1, 3) == L"1:21:27"
		   && TextAt (pg, load1, 4) == L"54:18:00 (40)" && TextAt (pg, load1, 5) == L"1:21:27"
		   && TextAt (pg, load1, 6) == L"168 -> 152" && TextAt (pg, load1, 7) == L"12.25 -> 11.10",
		   "a load: before -> after, saved a part and a batch of 40, the running total, flips, insert cost");
	Check (RowOf (pg, L"    why op 12 feed (0.3 -> 0.25): chatter on the shoulder", load1) > load1, "its reason under it");
	Check (TextAt (pg, load2, 3) == L"0:10:00" && TextAt (pg, load2, 5) == L"1:31:27"
		   && RowOf (pg, L"    undone 2026-10-11 10:30", load2) == load2 + 2, "the second load: its saving, then undone");
	Check (TextAt (pg, undo, 3) == L"-0:10:00" && TextAt (pg, undo, 5) == L"1:21:27", "the undo takes its saving back from the total");
	Check (TextAt (pg, dump, 2) == L"26:32:07" && TextAt (pg, dump, 6) == L"168" && TextAt (pg, dump, 7) == L"12.25",
		   "the dump: the part's figures");
	bool grey = false;
	for (const auto &row : pg.rows)
		for (const Cell &c : row)
			grey = grey || c.look == Cell::Auto;
	Check (!grey, "nothing on it drawn as a cell to type in or a grey one");

	const Xlsx::Sheet::Page empty = History::Page ({}, L"new.mcam", L"", L"");
	Check (RowOf (empty, L"No loads yet.") > 0 && RowOf (empty, L"Nothing measured on the machine yet") > 0, "a new part says so");
	}

	// ---- The workbook: the History page after the Summary, the main sheet still found.
	{
	std::vector<std::vector<std::wstring>> rows;
	std::vector<size_t> sheetRow;
	std::wstring why;
	Check (Xlsx::ReadSheet (wbFile, rows, sheetRow, why) && rows.size () == 5 && rows[1][0] == L"2",
		   "the load reads 'Lathe params', not the History");
	Xlsx::Grid hist;
	Check (Xlsx::ReadGrid (wbFile, L"History", hist, why) && hist[1][0] == L"History - SAMPLE.mcam", "the History sheet reads back");
	std::ifstream in (wbFile, std::ios::binary);
	const std::string bytes ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
	std::map<std::string, std::string> parts;
	Check (Xlsx::Unzip (bytes, parts, why), "unzips");
	const std::string &wb = parts["xl/workbook.xml"];
	const size_t sm = wb.find ("<sheet name=\"Summary\""), hi = wb.find ("<sheet name=\"History\""), mn = wb.find ("<sheet name=\"Lathe params\"");
	Check (sm < hi && hi < mn, "Summary, History, then the main sheet");
	Check (wb.find ("_xlnm._FilterDatabase\" localSheetId=\"2\"") != std::string::npos, "the filter names the main sheet at its new place");
	Check (wb.find ("<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"1\">'History'!$") != std::string::npos,
		   "the History's heading row printed on every page");
	Check (parts.count ("xl/worksheets/page1.xml") && parts["xl/worksheets/page1.xml"].find ("orientation=\"landscape\" fitToWidth=\"1\"") != std::string::npos
		   && parts["xl/worksheets/page1.xml"].find ("tabSelected") == std::string::npos,
		   "printed landscape, one page wide; the Summary is still the page that opens");
	Check (parts["[Content_Types].xml"].find ("/xl/worksheets/page1.xml") != std::string::npos
		   && parts["xl/_rels/workbook.xml.rels"].find ("Target=\"worksheets/page1.xml\"") != std::string::npos, "in the package's lists");
	}

	if (failed)
		{
		std::printf ("history_test: %d FAILED\n", failed);
		return 1;
		}
	std::puts ("history_test: all checks passed");
	return 0;
	}
