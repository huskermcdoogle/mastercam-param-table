//
// load_test.cpp - the load's preview and undo, outside Mastercam: the impact
// of the edits (Impact), the tick boxes (PreviewTicks), and reading the log
// back to decide what an undo restores (Undo).
//
#include "../src/Impact.h"
#include "../src/Preview.h"
#include "../src/Undo.h"

#include <cstdio>

namespace
	{
	int gFailed = 0;

	void Check (bool ok, const char *what)
		{
		if (!ok)
			{
			std::printf ("FAIL  %s\n", what);
			++gFailed;
			}
		}

	Preview::Line L (Preview::Line::Kind k, const wchar_t *text, Preview::Line::Box box,
					 const wchar_t *link = L"")
		{
		Preview::Line l;
		l.kind = k;
		l.text = text;
		l.box = box;
		l.link = link;
		return l;
		}

	Plan::Schema Rough ()
		{
		Plan::Schema s;
		s.type = L"ROUGH";
		Plan::Col feed;  feed.name = L"feed";       feed.type = Plan::Type::Double;
		Plan::Col mode;  mode.name = L"feed_mode";  mode.type = Plan::Type::Text;
		mode.choices = { L"per rev", L"per min" };
		Plan::Col text;  text.name = L"manual_text"; text.type = Plan::Type::Text;
		Plan::Col cool;  cool.name = L"coolant";    cool.type = Plan::Type::Text;
		s.cols = { feed, mode, text, cool };
		return s;
		}

	Plan::Current Op (long op, const wchar_t *feed, const wchar_t *mode, const wchar_t *text,
					  const wchar_t *cool)
		{
		Plan::Current c;
		c.op = op;
		c.type = L"ROUGH";
		c.values = { feed, mode, text, cool };
		return c;
		}
	}

int main ()
	{
	// ---- IMPACT: time as dumped vs the live estimate, flips from the Dumped
	// sheet, and the total following what is applied.
	{
	const std::vector<Csv::Row> sheet = {
		{ L"op_idn", L"type", L"cycle_time_raw", L"est_seconds", L"flips", L"flips_part" },
		{ L"1", L"ROUGH", L"3600", L"3000", L"2", L"5" },
		{ L"2", L"FINISH", L"600", L"600", L"0", L"1" },
		{ L"", L"", L"", L"", L"", L"" },					// a blank row: no op
		{ L"3", L"FACE", L"", L"", L"", L"" } };				// no figures at all
	const std::vector<Csv::Row> dumped = {
		{ L"op_idn", L"type", L"flips_part" },
		{ L"1", L"ROUGH", L"6" },
		{ L"2", L"FINISH", L"1" } };

	const auto ops = Impact::Read (sheet, dumped);
	Check (ops.size () == 3, "impact: one entry per op, blank rows left out");
	Check (ops.at (1).was == 3600 && ops.at (1).now == 3000, "impact: time as dumped and now");
	Check (ops.at (1).flipsWas == 6 && ops.at (1).flipsNow == 5,
		   "impact: flips as dumped come from the Dumped sheet's flips_part, not flips");
	Check (std::isnan (ops.at (3).was), "impact: an op without figures says so");

	Impact::Total all = Impact::Sum (ops, { 1, 2 });
	Check (all.was == 4200 && all.now == 3600 && all.timed == 2, "impact: total with every edit");
	Check (all.flipsWas == 7 && all.flipsNow == 6, "impact: flips total with every edit");
	Check (Impact::TotalText (all) == L"Cycle time 1:10:00 -> 1:00:00 (-0:10:00)  ·  flips 7 -> 6",
		   "impact: the summary line");

	Impact::Total none = Impact::Sum (ops, {});
	Check (none.now == none.was && none.flipsNow == none.flipsWas,
		   "impact: an op whose changes are all left out counts as dumped");
	Check (Impact::TotalText (none) == L"Cycle time 1:10:00, unchanged  ·  flips 7, unchanged",
		   "impact: nothing applied reads as unchanged");

	Check (Impact::OpText (ops.at (1)) == L"1:00:00 -> 0:50:00 (-0:10:00)  ·  flips 6 -> 5",
		   "impact: one op's line");
	Check (Impact::OpText (ops.at (2)).empty (), "impact: an op the edits do not move says nothing");

	// Without a Dumped sheet, flips as dumped is the flips column.
	const auto bare = Impact::Read (sheet, {});
	Check (bare.at (1).flipsWas == 2, "impact: no Dumped sheet - flips as dumped from flips");

	// A CSV of some other kind: no figures, no summary.
	const auto csv = Impact::Read ({ { L"op_idn", L"type", L"feed" }, { L"1", L"ROUGH", L"0.3" } }, {});
	Check (Impact::TotalText (Impact::Sum (csv, { 1 })).empty (), "impact: a sheet without figures has no line");

	Check (Impact::Hms (95527) == L"26:32:07", "hms: hours past 24 are not wrapped");
	Check (Impact::Hms (252.4) == L"0:04:12", "hms: under an hour, rounded to the second");
	Check (Impact::Change (-4887) == L"-1:21:27", "change: a saving");
	Check (Impact::Change (-0.4) == L"+0:00:00", "change: under half a second is no change, never -0:00:00");
	Check (Impact::Tone (100, 90) == -1 && Impact::Tone (90, 100) == 1 && Impact::Tone (90, 90.3) == 0,
		   "tone: better, worse, same to the second");
	}

	// ---- TICK BOXES: an operation's box takes its changes; linked changes go
	// together; the operation's box follows.
	{
	using Line = Preview::Line;
	std::vector<Line> ls = {
		L (Line::Section, L"Changes", Line::NoBox),
		L (Line::Op, L"op 1", Line::Ticked),
		L (Line::Change, L"feed", Line::Ticked, L"feed"),
		L (Line::Change, L"feed_mode", Line::Ticked, L"feed"),
		L (Line::Change, L"step", Line::Ticked, L"step"),
		L (Line::Op, L"op 2", Line::Ticked),
		L (Line::Change, L"feed", Line::Ticked, L"feed"),
		L (Line::Section, L"Skipped", Line::NoBox),
		L (Line::Note, L"x", Line::NoBox) };

	Check (Preview::CountTicked (ls) == 4, "ticks: all changes start ticked");
	Check (Preview::OpOf (ls, 4) == 1 && Preview::OpOf (ls, 6) == 5 && Preview::OpOf (ls, 8) == ls.size (),
		   "ticks: each change knows its operation");

	Preview::Toggle (ls, 2);
	Check (ls[2].box == Line::Unticked && ls[3].box == Line::Unticked,
		   "ticks: unticking a feed unticks its feed mode too");
	Check (ls[4].box == Line::Ticked, "ticks: an unlinked change is untouched");
	Check (ls[6].box == Line::Ticked, "ticks: the same column on ANOTHER operation is untouched");
	Check (ls[1].box == Line::Mixed, "ticks: an operation partly ticked shows mixed");

	Preview::Toggle (ls, 1);
	Check (ls[1].box == Line::Ticked && ls[2].box == Line::Ticked && ls[3].box == Line::Ticked,
		   "ticks: clicking a mixed operation ticks all of it");
	Preview::Toggle (ls, 1);
	Check (ls[1].box == Line::Unticked && ls[4].box == Line::Unticked && ls[6].box == Line::Ticked,
		   "ticks: clicking it again unticks all of it - and only it");
	Check (Preview::CountTicked (ls) == 1, "ticks: the count follows");

	Preview::Toggle (ls, 4);
	Check (ls[1].box == Line::Mixed, "ticks: ticking one change back makes the operation mixed");
	Preview::Toggle (ls, 0);
	Preview::Toggle (ls, 8);
	Check (ls[0].box == Line::NoBox && ls[8].box == Line::NoBox, "ticks: lines without a box ignore clicks");
	}

	// ---- LINKED COLUMNS
	{
	Check (Plan::LinkOf (L"feed_mode") == Plan::LinkOf (L"feed"), "link: feed and its mode");
	Check (Plan::LinkOf (L"speed_mode") == L"speed" && Plan::LinkOf (L"spindle_dir") == L"speed",
		   "link: speed, CSS/RPM and direction");
	Check (Plan::LinkOf (L"insp_time_on") == L"insp_time", "link: a switch and its value");
	Check (Plan::LinkOf (L"step") == L"step" && Plan::LinkOf (L"max_ss") == L"max_ss",
		   "link: anything else is its own");
	}

	// ---- THE LOG LINES: written and read by the same code.
	{
	Check (Undo::ChangeLine (12, L"ROUGH", L"feed", L"0.3", L"0.25") == L"op 12 ROUGH  feed  0.3 -> 0.25",
		   "log: the change line keeps its old form");
	Check (Undo::OneLine (L"M00\r\nM01") == L"M00 ↵ M01", "log: line breaks shown on one line");
	Check (Undo::FromOneLine (L"M00 ↵ M01", L"") == L"M00\r\nM01", "log: line breaks back, CR LF");
	Check (Undo::FromOneLine (L"M00 ↵ M01", L"a\nb") == L"M00\nM01",
		   "log: line breaks back as the value now has them");
	const auto sp = Undo::Splits (L"(X -> Z) -> (Z)");
	Check (sp.size () == 2 && sp[0].first == L"(X" && sp[1].first == L"(X -> Z)" && sp[1].second == L"(Z)",
		   "log: a value holding an arrow offers every split");
	const auto empty = Undo::Splits (L" -> M00");
	Check (!empty.empty () && empty[0].first.empty () && empty[0].second == L"M00",
		   "log: an empty old value");
	}

	// ---- FINDING THE LAST LOAD of this part in a log shared by the folder.
	const std::wstring log =
		L"﻿2026-10-01 09:00:00  op 5 ROUGH  feed  0.2 -> 0.3\r\n"			// old version: no begin line
		L"2026-10-01 09:00:00  load: 1 operation(s) written, 0 failed\r\n"
		L"2026-10-02 10:00:00  REFUSED shaft_1.xlsx row 4: \"feed\": \"x\" is not a number\r\n"
		L"2026-10-02 10:00:01  load begin: shaft.mcam  <-  shaft_1.xlsx\r\n"
		L"2026-10-02 10:00:01  op 12 ROUGH  feed  0.3 -> 0.25\r\n"
		L"2026-10-02 10:00:01  load: 1 operation(s) written, 0 failed\r\n"
		L"2026-10-03 11:00:00  load begin: Shaft.MCAM  <-  shaft_2.xlsx, shaft_3.xlsx\r\n"
		L"2026-10-03 11:00:00  left out: op 12 ROUGH  coolant  Off -> Flood\r\n"
		L"2026-10-03 11:00:00  op 12 ROUGH  feed  0.25 -> 0.2\r\n"
		L"2026-10-03 11:00:00  op 12 ROUGH  feed_mode  per rev -> per min\r\n"
		L"2026-10-03 11:00:00  op 14 ROUGH  feed  0.4 -> 0.35\r\n"
		L"2026-10-03 11:00:00  op 14 FAILED writing feed\r\n"
		L"2026-10-03 11:00:00  op 15 ROUGH  manual_text  M00 ↵ M01 -> M00 ↵ M05\r\n"
		L"2026-10-03 11:00:00  op 16 ROUGH  feed  0.1 -> 0.15\r\n"
		L"2026-10-03 11:00:00  op 16 ROUGH  feed  0.1 -> 0.12\r\n"			// a second sheet, same op
		L"2026-10-03 11:00:00  op 17 MILL DRILL  feed  1 -> 2\r\n"
		L"2026-10-03 11:00:01  load: 4 operation(s) written, 1 failed\r\n"
		L"2026-10-04 08:00:00  load begin: other.mcam  <-  other.xlsx\r\n"
		L"2026-10-04 08:00:00  op 99 ROUGH  feed  1 -> 2\r\n"
		L"2026-10-04 08:00:00  load: 1 operation(s) written, 0 failed\r\n"
		L"2026-10-05 08:00:00  undo begin: shaft.mcam  (load of 2026-10-03 11:00:00)\r\n"
		L"2026-10-05 08:00:00  undo op 12 ROUGH  feed  0.2 -> 0.25\r\n"
		L"2026-10-05 08:00:00  undo: 1 operation(s) restored, 0 failed\r\n";

	const Undo::Last last = Undo::FindLast (log, L"shaft.mcam");
	Check (last.found && last.stamp == L"2026-10-03 11:00:00", "find: the newest load of this part, by name in any case");
	Check (last.files == L"shaft_2.xlsx, shaft_3.xlsx", "find: the sheets it loaded");
	Check (last.end == L"load: 4 operation(s) written, 1 failed", "find: its closing line");
	Check (last.entries.size () == 6, "find: its values - left-out and FAILED lines are not values");
	if (last.entries.size () == 6)
		{
		Check (last.entries[0].op == 12 && last.entries[0].type == L"ROUGH" && last.entries[0].column == L"feed"
			   && last.entries[0].both == L"0.25 -> 0.2", "find: an entry's parts");
		Check (last.entries[4].both == L"0.1 -> 0.12", "find: the same value twice - the later write is kept");
		Check (last.entries[5].type == L"MILL DRILL", "find: a type with a space in it");
		}
	Check (!Undo::FindLast (log, L"other.mcam").entries.empty (), "find: another part's load is its own");
	const Undo::Last none = Undo::FindLast (log, L"nothing.mcam");
	Check (!none.found && none.olderLoads, "find: no load of the part, and the log has older loads without a part");
	Check (Undo::BeginLine (L"shaft.mcam", L"a.xlsx") == L"load begin: shaft.mcam  <-  a.xlsx",
		   "log: the begin line FindLast reads");

	// ---- WHAT AN UNDO RESTORES.
	{
	const Plan::Schema rough = Rough ();
	Undo::Kind k;
	k.schema = &rough;
	k.now = {
		Op (12, L"0.2", L"per min", L"", L"Off"),		// as the load left it: restored
		Op (14, L"0.4", L"per rev", L"", L"Off"),		// the write failed: already as before
		Op (15, L"", L"per rev", L"M00\r\nM05", L"Off"),	// manual text, line breaks back
		Op (16, L"0.3", L"per rev", L"", L"Off") };	// changed since: left alone
	const Undo::Result r = Undo::Make (last, { k });

	Check (r.restores.size () == 3, "undo: three values to restore");
	if (r.restores.size () == 3)
		{
		Check (r.restores[0].change.op == 12 && r.restores[0].change.name == L"feed"
			   && r.restores[0].change.from == L"0.2" && r.restores[0].change.to == L"0.25",
			   "undo: a value still as the load left it goes back");
		Check (r.restores[1].change.name == L"feed_mode" && r.restores[1].change.to == L"per rev",
			   "undo: its feed mode too");
		Check (r.restores[2].change.op == 15 && r.restores[2].change.to == L"M00\r\nM01",
			   "undo: a manual entry's line breaks come back");
		}
	Check (r.already.size () == 1 && r.already[0].op == 14, "undo: a value already back is left alone");

	bool sawChanged = false, sawGone = false;
	for (const Undo::Skip &s : r.skipped)
		{
		if (s.op == 16 && s.column == L"feed" && s.why.find (L"changed since") != std::wstring::npos)
			sawChanged = true;
		if (s.op == 17 && s.column.empty ())
			sawGone = true;
		}
	Check (sawChanged, "undo: a value changed since the load is skipped and says so");
	Check (sawGone, "undo: an operation no longer in the part (or of an unknown kind) is skipped");

	// One value changed since holds back the whole operation.
	Undo::Last two;
	two.found = true;
	Undo::Entry a, b;
	a.op = 12; a.type = L"ROUGH"; a.column = L"feed"; a.both = L"0.25 -> 0.2";
	b.op = 12; b.type = L"ROUGH"; b.column = L"feed_mode"; b.both = L"per rev -> per min";
	two.entries = { a, b };
	Undo::Kind k2;
	k2.schema = &rough;
	k2.now = { Op (12, L"0.18", L"per min", L"", L"Off") };
	const Undo::Result r2 = Undo::Make (two, { k2 });
	Check (r2.restores.empty () && r2.skipped.size () == 2,
		   "undo: one value changed since holds back its whole operation");

	// Numbers compare as numbers, choices and text ignoring case.
	Undo::Entry c, d;
	c.op = 12; c.type = L"ROUGH"; c.column = L"feed"; c.both = L"0.25 -> 0.2";
	d.op = 12; d.type = L"ROUGH"; d.column = L"coolant"; d.both = L"Off -> flood";
	two.entries = { c, d };
	k2.now = { Op (12, L"0.2000000000001", L"per min", L"", L"Flood") };
	const Undo::Result r3 = Undo::Make (two, { k2 });
	Check (r3.restores.size () == 2 && r3.skipped.empty (),
		   "undo: \"0.2\" written reads back as 0.2000000000001, \"flood\" as Flood - still the load's");

	// A text value holding an arrow: the split that matches the operation wins.
	Undo::Entry e;
	e.op = 12; e.type = L"ROUGH"; e.column = L"manual_text"; e.both = L"(X -> Z) -> (Z)";
	two.entries = { e };
	k2.now = { Op (12, L"0.2", L"per min", L"(Z)", L"Off") };
	const Undo::Result r4 = Undo::Make (two, { k2 });
	Check (r4.restores.size () == 1 && r4.restores[0].change.to == L"(X -> Z)",
		   "undo: an arrow inside a value does not confuse old and new");
	}

	if (gFailed != 0)
		{
		std::printf ("load_test: %d check(s) FAILED\n", gFailed);
		return 1;
		}
	std::printf ("load_test: all checks passed\n");
	return 0;
	}
