// The batch dump's shop workbook, without Mastercam: two parts' dumps laid out as
// src\Dump.cpp writes them (the main sheet's cached values, the Tools page, its
// inserts table) read back, compared and grouped - speeds and feeds worked out
// as the dump's MRR does, metric turned to inch, edge times as the Program check
// works them out, medians and ranges per insert - then written as a workbook and
// read back by the add-in's own reader. Also the folder search.
//
// Usage: shop_test <out dir>   (writes shop_sample.xlsx there, and a scratch folder)
#include "../src/Shop.h"
#include "../src/Xlsx.h"

#include <algorithm>
#include <cmath>
#include <cstdio>
#include <ctime>
#include <filesystem>
#include <fstream>
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

	bool Near (double a, double b, double tol = 1e-6)
		{
		return std::fabs (a - b) <= tol * (std::max) (1.0, std::fabs (b));
		}

	Xlsx::Sheet::FreeCell Cell (const std::wstring &t, bool head = false)
		{
		Xlsx::Sheet::FreeCell c;
		c.text = t;
		c.head = head;
		return c;
		}

	Xlsx::Sheet::ToolRow Tool (const wchar_t *number, const wchar_t *usedBy, const wchar_t *insert)
		{
		Xlsx::Sheet::ToolRow t;
		t.number = number;
		t.name = L"tool";
		t.usedBy = usedBy;
		// As the dump leaves it: Insert, Flips / part, Cut time / part, Longest ... (by heading).
		t.extra = { Cell (insert), Cell (L"0"), Cell (L"0:00:00"), Cell (L"") };
		return t;
		}

	/// The inserts table under the tools: its heading row, then per insert
	/// (name, edges, flips, cost, parts per edge), then the spare rows.
	std::vector<std::vector<Xlsx::Sheet::FreeCell>> Inserts (
		const std::vector<std::vector<std::wstring>> &rows)
		{
		std::vector<std::vector<Xlsx::Sheet::FreeCell>> out;
		std::vector<Xlsx::Sheet::FreeCell> head;
		for (const wchar_t *h : { L"Inserts", L"Insert", L"Used by", L"Edges per insert", L"Flips / part", L"Inserts / part",
								  L"Cost per insert", L"Insert cost / part", L"Parts per edge", L"Usual edge time" })
			head.push_back (Cell (h, true));
		out.push_back (head);
		for (const auto &r : rows)
			out.push_back ({ Cell (L""), Cell (r[0]), Cell (L"T?"), Cell (r[1]), Cell (r[2]), Cell (L""), Cell (r[3]), Cell (L""),
							 Cell (r[4]), Cell (L"") });
		for (int spare = 0; spare < 4; ++spare)
			out.push_back ({ Cell (L""), Cell (L""), Cell (L""), Cell (L""), Cell (L""), Cell (L""), Cell (L""), Cell (L""),
							 Cell (L""), Cell (L"") });
		return out;
		}

	const Shop::Op *OpOf (const Shop::Part &p, const wchar_t *op)
		{
		for (const Shop::Op &o : p.ops)
			if (o.op == op)
				return &o;
		return nullptr;
		}
	}

int main (int argc, char **argv)
	{
	const std::string dir = argc > 1 ? argv[1] : ".";
	const double pi = 3.14159265358979323846;

	// ---- Medians, ranges, edge times.
	{
	const Shop::Spread odd = Shop::SpreadOf ({ 3, 1, 2 });
	Check (odd.n == 3 && odd.min == 1 && odd.median == 2 && odd.max == 3, "spread of an odd count: the middle one");
	const Shop::Spread even = Shop::SpreadOf ({ 4, 1, 3, 2 });
	Check (even.n == 4 && even.median == 2.5 && even.min == 1 && even.max == 4, "spread of an even count: the mean of the two middle");
	const Shop::Spread gaps = Shop::SpreadOf ({ Shop::kNone, 5, Shop::kNone });
	Check (gaps.n == 1 && gaps.median == 5 && gaps.min == 5, "unknown values are left out");
	const Shop::Spread none = Shop::SpreadOf ({});
	Check (none.n == 0 && std::isnan (none.median) && std::isnan (none.min), "nothing: no median");

	Check (Shop::EdgeSeconds (600, 3, 0, 0) == 200, "edge time: cut over flips");
	Check (Shop::EdgeSeconds (600, 3, 250, 0) == 250, "edge time: the longest between flips when longer");
	Check (Shop::EdgeSeconds (600, 3, 100, 0) == 200, "edge time: not a shorter longest");
	Check (Shop::EdgeSeconds (300, 0.5, 0, 4) == 1200, "edge time: an edge that outlasts the part - cut x parts per edge");
	Check (std::isnan (Shop::EdgeSeconds (300, 0.5, 0, 0)), "edge time: outlasts the part, nobody said by how much - unknown");
	Check (std::isnan (Shop::EdgeSeconds (0, 2, 0, 0)), "edge time: no cutting - unknown");
	}

	// ---- Part A (inch): a dump as written.
	Xlsx::Sheet a;
	//           0          1        2        3           4         5         6              7         8             9
	a.rows = { { L"op_idn", L"type", L"tool", L"comment", L"changes", L"units", L"needs_regen", L"speed", L"speed_mode", L"max_ss",
	//           10       11           12       13        14             15          16             17                 18
				 L"feed", L"feed_mode", L"step", L"n_cuts", L"rough_step", L"cut_dia", L"est_seconds", L"cut_seconds_est", L"flips_part",
	//           19               20       21          22             23           24                 25              26
				 L"flip_longest", L"mrr", L"air_pct", L"xf_copies", L"do_rough", L"pt_rough_speed", L"pt_rough_css",
				 L"pt_rough_feed_axial", L"pt_rough_axial_type" },
			   { L"1", L"ROUGH", L"1", L"Rough OD", L"0", L"in", L"no", L"500", L"CSS", L"3000", L"0.012", L"per rev", L"0.1", L"",
				 L"", L"2", L"300", L"240", L"2", L"2:30", L"7.2", L"10", L"", L"", L"", L"", L"", L"" },
			   { L"2", L"FINISH", L"3", L"Finish OD", L"0", L"in", L"no", L"600", L"CSS", L"3000", L"0.006", L"per rev", L"0.02", L"1",
				 L"", L"1.5", L"120", L"100", L"0.5", L"", L"", L"", L"", L"", L"", L"", L"", L"" },
			   { L"3", L"GROOVE", L"3", L"Groove", L"0", L"in", L"yes", L"1000", L"RPM", L"3000", L"2", L"per min", L"", L"",
				 L"0.05", L"1", L"60", L"50", L"0.5", L"", L"", L"", L"1", L"", L"", L"", L"", L"" },
			   { L"4", L"PRIME", L"5", L"Prime", L"0", L"in", L"no", L"0", L"RPM", L"0", L"0", L"per rev", L"0.08", L"",
				 L"", L"3", L"200", L"180", L"0.25", L"", L"", L"", L"", L"1", L"800", L"1", L"0.02", L"per rev" },
			   { L"5", L"MANUAL", L"0", L"M01", L"0", L"", L"no", L"", L"", L"", L"", L"", L"", L"",
				 L"", L"", L"", L"", L"", L"", L"", L"", L"", L"", L"", L"", L"", L"" },
			   { L"6", L"TRANSFORM", L"0", L"Copies", L"0", L"", L"no", L"", L"", L"", L"", L"", L"", L"",
				 L"", L"", L"400", L"300", L"1", L"", L"", L"", L"", L"", L"", L"", L"", L"" } };
	a.toolExtraHeads = { L"Insert", L"Flips / part", L"Cut time / part", L"Longest between flips" };
	a.tools = { Tool (L"1", L"op 1", L"C 80° IC 0.5 r0.031"), Tool (L"3", L"op 2, 3", L"CNMG432"),
				Tool (L"5", L"op 4", L"PrimeTurning B r0.016") };
	a.toolsAfter = Inserts ({ { L"C 80° IC 0.5 r0.031", L"2", L"2", L"12.5", L"" },
							  { L"CNMG432", L"4", L"1.5", L"", L"" },
							  { L"PrimeTurning B r0.016", L"", L"0.25", L"", L"4" } });

	Shop::Part pa;
	pa.name = L"Shaft";
	pa.file = L"Shaft.mcam";
	pa.ok = true;
	pa.material = L"4140";
	Shop::Read (a, pa);
	Check (pa.ops.size () == 6, "every op row read");
	Check (pa.units == L"in", "the part's units from its ops");
	Check (pa.regen == 1, "ops needing regen counted");
	Check (Near (pa.cycleSeconds, 300 + 120 + 60 + 200 + 400), "cycle time: est_seconds added up, as the Summary does");
	Check (Near (pa.cutSeconds, 240 + 100 + 50 + 180 + 300), "cutting time added up");
	Check (Near (pa.flips, 2 + 0.5 + 0.5 + 0.25 + 1), "flips per part added up (the transform's copies are its own row)");
	Check (pa.inserts == 3 && pa.costed == 1, "inserts used, and how many have a cost");
	Check (Near (pa.insertCost, 2.0 / 2 * 12.5), "insert cost per part: flips / edges x cost, over the costed ones");

	if (const Shop::Op *o = OpOf (pa, L"1"))
		{
		Check (o->tool == L"1" && o->insert == L"C 80° IC 0.5 r0.031", "an op's tool and insert from the Tools page");
		Check (Near (o->sfm, 500), "CSS under max RPM: the surface speed is the speed");
		Check (Near (o->feedIpr, 0.012), "feed per rev as is");
		Check (Near (o->depthIn, 0.1), "rough: depth is step");
		Check (Near (o->edgeSeconds, 150), "edge time: the longest between flips (2:30) beats 240 s / 2 flips");
		Check (Near (o->mrr, 7.2) && o->mrrUnits == L"in3/min", "MRR and its units");
		Check (Near (o->airPct, 10), "air share");
		}
	else
		Check (false, "op 1 read");
	if (const Shop::Op *o = OpOf (pa, L"2"))
		{
		Check (o->depth.empty () && std::isnan (o->depthIn), "a one-pass finish has no depth of cut");
		// T3: op 2 (100 s, 0.5 flips) and op 3 (50 s, 0.5 flips) cut twice - a transform copies it.
		Check (Near (o->edgeSeconds, (100 + 50 * 2) / (0.5 + 0.5 * 2)), "edge time per TOOL: its ops together, transform copies counted");
		}
	else
		Check (false, "op 2 read");
	if (const Shop::Op *o = OpOf (pa, L"3"))
		{
		Check (Near (o->sfm, 1000 * pi * 1 / 12), "RPM: surface speed at the cutting diameter");
		Check (Near (o->feedIpr, 2.0 / 1000), "per-minute feed over the RPM");
		Check (Near (o->depthIn, 0.05), "groove: depth is rough step");
		Check (o->needsRegen == L"yes", "needs regen carried");
		}
	else
		Check (false, "op 3 read");
	if (const Shop::Op *o = OpOf (pa, L"4"))
		{
		Check (o->speed == L"800" && o->speedMode == L"CSS" && o->feed == L"0.02", "PrimeTurning: its own rough pass's speed and feed");
		Check (Near (o->sfm, 800) && Near (o->feedIpr, 0.02) && Near (o->depthIn, 0.08), "PrimeTurning in inch units");
		Check (Near (o->edgeSeconds, 180 * 4), "edge that outlasts the part: cut time x parts per edge");
		}
	else
		Check (false, "op 4 read");
	if (const Shop::Op *o = OpOf (pa, L"5"))
		Check (o->tool.empty () && o->insert.empty () && std::isnan (o->edgeSeconds), "manual entry: no tool, no insert");
	else
		Check (false, "op 5 read");

	// The cap: CSS held to max RPM at a small diameter.
	{
	Xlsx::Sheet c;
	c.rows = { { L"op_idn", L"type", L"units", L"speed", L"speed_mode", L"max_ss", L"feed", L"feed_mode", L"cut_dia" },
			   { L"1", L"FINISH", L"in", L"600", L"CSS", L"1000", L"10", L"per min", L"0.5" } };
	Shop::Part pc;
	Shop::Read (c, pc);
	const double rpm = 1000;			// 600 SFM at 0.5 dia wants 4584 RPM
	Check (pc.ops.size () == 1 && Near (pc.ops[0].sfm, rpm * pi * 0.5 / 12), "CSS capped by max_ss: the speed at the cap");
	Check (pc.ops.size () == 1 && Near (pc.ops[0].feedIpr, 10.0 / rpm), "per-minute feed over the capped RPM");
	Check (pc.ops.size () == 1 && pc.ops[0].tool.empty () && pc.ops[0].insert.empty (), "no Tools page: read, with no tool or insert");
	}

	// ---- Part B (metric): the same insert by another spelling, and one of its own.
	Xlsx::Sheet b;
	b.rows = { { L"op_idn", L"type", L"tool", L"units", L"speed", L"speed_mode", L"max_ss", L"feed", L"feed_mode", L"n_cuts", L"step",
				 L"cut_dia", L"est_seconds", L"cut_seconds_est", L"flips_part" },
			   { L"10", L"FINISH", L"7", L"mm", L"200", L"CSS", L"4000", L"0.3", L"per rev", L"2", L"2.5", L"50", L"90", L"80", L"1" },
			   { L"11", L"ROUGH", L"8", L"mm", L"180", L"CSS", L"4000", L"0.35", L"per rev", L"", L"3", L"60", L"500", L"450", L"3" } };
	b.toolExtraHeads = { L"Insert" };
	b.tools = { Tool (L"7", L"op 10", L"cnmg432"), Tool (L"8", L"op 11", L"WNMG080408") };
	b.toolsAfter = Inserts ({ { L"cnmg432", L"4", L"1", L"8", L"" }, { L"WNMG080408", L"6", L"3", L"9", L"" } });
	Shop::Part pb;
	pb.name = L"Flange";
	pb.file = L"Metric\\Flange.mcam";
	pb.ok = true;
	pb.material = L"4140";
	Shop::Read (b, pb);
	if (const Shop::Op *o = OpOf (pb, L"10"))
		{
		Check (Near (o->sfm, 200 * 1000 / 304.8), "metric: m/min to feet per minute");
		Check (Near (o->feedIpr, 0.3 / 25.4), "metric: mm/rev to inches per rev");
		Check (Near (o->depthIn, 2.5 / 25.4), "a finish of several passes: depth is step, mm to inches");
		Check (o->speed == L"200" && o->feed == L"0.3", "the op's own values kept as dumped");
		Check (o->mrrUnits.empty (), "no MRR, no units");
		}
	else
		Check (false, "op 10 read");
	Check (pb.costed == 2 && Near (pb.insertCost, 1.0 / 4 * 8 + 3.0 / 6 * 9), "insert cost per part, every insert costed");

	// ---- A part that would not open.
	Shop::Part pf;
	pf.name = L"Broken";
	pf.file = L"Broken.mcam";
	pf.ok = false;
	pf.why = L"Mastercam could not open it";

	const std::vector<Shop::Part> parts = { pa, pb, pf };

	// ---- Per insert: grouped by insert (any case), kind of op and material.
	{
	const std::vector<Shop::InsertRow> rows = Shop::ByInsert (parts);
	const Shop::InsertRow *fin = nullptr;
	size_t cnmgRows = 0;
	for (const Shop::InsertRow &r : rows)
		{
		if (r.insert == L"CNMG432" || r.insert == L"cnmg432")
			++cnmgRows;
		if ((r.insert == L"CNMG432" || r.insert == L"cnmg432") && r.opKind == L"FINISH")
			fin = &r;
		}
	Check (cnmgRows == 2, "CNMG432 at two kinds of op: two rows (FINISH, GROOVE)");
	Check (rows.size () == 5, "five insert / kind rows (C80 ROUGH, CNMG FINISH, CNMG GROOVE, Prime PRIME, WNMG ROUGH)");
	if (fin != nullptr)
		{
		Check (fin->nOps == 2 && fin->parts.size () == 2, "the same insert across two parts, spelt two ways: one row");
		Check (fin->insert == L"CNMG432", "the first spelling seen is shown");
		Check (fin->material == L"4140", "its material");
		Check (Near (fin->sfm.median, (600 + 200 * 1000 / 304.8) / 2) && Near (fin->sfm.min, 600), "SFM range and median, metric converted");
		Check (Near (fin->feed.median, (0.006 + 0.3 / 25.4) / 2), "feed median in inches per rev");
		Check (fin->depth.n == 1 && Near (fin->depth.median, 2.5 / 25.4), "depth: only the op that has one");
		Check (fin->edge.n == 2 && Near (fin->edge.min, 80) && Near (fin->edge.max, 200 / 1.5), "edge times, once per tool");
		Check (Near (fin->flips, 0.5 + 1), "flips added over its ops");
		Check (fin->tools == L"Shaft: T3; Flange: T7", "its tools, per part");
		}
	else
		Check (false, "the CNMG432 FINISH row");
	bool broken = false;
	for (const Shop::InsertRow &r : rows)
		for (const std::wstring &p : r.parts)
			broken = broken || p == L"Broken";
	Check (!broken, "a part that was not dumped adds nothing");
	}

	// ---- The workbook: built, read back by the add-in's own reader.
	{
	const std::vector<Xlsx::Table> tables = Shop::Tables (parts, L"Shop compare - test");
	Check (tables.size () == 4 && tables[0].name == L"Parts" && tables[1].name == L"Inserts" && tables[2].name == L"Ops"
		   && tables[3].name == L"About", "sheets: Parts, Inserts, Ops, About");
	const std::string bytes = Xlsx::BuildBook (tables);
	{
	std::ofstream f (dir + "\\shop_sample.xlsx", std::ios::binary | std::ios::trunc);
	f.write (bytes.data (), static_cast<std::streamsize> (bytes.size ()));
	}
	Xlsx::Grid g;
	std::wstring why;
	const bool read = Xlsx::ReadGridBytes (bytes, L"Inserts", g, why);
	Check (read, "the Inserts sheet reads back");
	auto colOf = [&g] (const wchar_t *name) -> long
		{
		for (const auto &c : g[1])
			if (c.second == name)
				return static_cast<long> (c.first);
		return -1;
		};
	bool all = true;
	for (const wchar_t *name : { L"insert", L"op_kind", L"material", L"sfm_median", L"feed_median", L"depth_median",
								 L"edge_time_median", L"n_ops", L"parts" })
		all = all && colOf (name) >= 0;
	Check (all, "the column names the standards library reads are all there");
	// Find the CNMG432 FINISH row and its median.
	bool found = false;
	for (const auto &r : g)
		if (r.first > 1 && r.second.count (static_cast<size_t> (colOf (L"insert")))
			&& r.second.at (static_cast<size_t> (colOf (L"insert"))) == L"CNMG432"
			&& r.second.at (static_cast<size_t> (colOf (L"op_kind"))) == L"FINISH")
			{
			found = true;
			double v = 0;
			Check (swscanf_s (r.second.at (static_cast<size_t> (colOf (L"sfm_median"))).c_str (), L"%lf", &v) == 1
				   && Near (v, std::round ((600 + 200 * 1000 / 304.8) / 2 * 10) / 10), "sfm_median written as a number, 1 place");
			Check (r.second.at (static_cast<size_t> (colOf (L"parts"))) == L"2", "parts: how many");
			}
	Check (found, "the CNMG432 FINISH row is on the sheet");

	Xlsx::Grid parts2;
	Check (Xlsx::ReadGridBytes (bytes, L"Parts", parts2, why) && parts2.size () == 4, "Parts: a heading row and three parts");
	bool notDumped = false;
	for (const auto &r : parts2)
		for (const auto &c : r.second)
			notDumped = notDumped || c.second == L"Mastercam could not open it";
	Check (notDumped, "the part that failed is listed, with why");

	Xlsx::Grid ops;
	Check (Xlsx::ReadGridBytes (bytes, L"Ops", ops, why) && ops.size () == 1 + 6 + 2, "Ops: every op of the parts dumped");
	Xlsx::Grid about;
	Check (Xlsx::ReadGridBytes (bytes, L"About", about, why) && about.size () > 20, "About explains the columns");
	}

	// ---- Same-named sheets are kept apart (Excel refuses a repeat).
	{
	Xlsx::Table t1, t2;
	t1.name = L"Ops";
	t2.name = L"ops";
	t1.rows = t2.rows = { { L"x" }, { L"1" } };
	Xlsx::Table t3;
	t3.name = L"a/b:c";
	t3.rows = { { L"x" } };
	const std::string bytes = Xlsx::BuildBook ({ t1, t2, t3 });
	Xlsx::Grid g;
	std::wstring why;
	Check (Xlsx::ReadGridBytes (bytes, L"a_b_c", g, why), "characters Excel refuses in a tab name become _");
	}

	// ---- The folder search.
	{
	namespace fs = std::filesystem;
	const fs::path root = fs::path (dir) / "shop_find";
	std::error_code ec;
	fs::remove_all (root, ec);
	fs::create_directories (root / "Sub", ec);
	for (const char *f : { "b.mcam", "A.MCAM", "notes.txt", "c.mcam-bak", "Sub/d.mcam" })
		std::ofstream (root / f) << "x";
	bool more = false;
	const auto top = Shop::FindParts (root, false, L".mcam", 0, more);
	Check (top.size () == 2 && top[0].filename () == L"A.MCAM" && top[1].filename () == L"b.mcam" && !more,
		   "this folder: the part files only, any case, in name order");
	const auto deep = Shop::FindParts (root, true, L".mcam", 0, more);
	Check (deep.size () == 3, "with subfolders: theirs too");
	const auto capped = Shop::FindParts (root, true, L".mcam", 1, more);
	Check (capped.size () == 1 && more, "a limit stops the search, and says there were more");
	Check (Shop::FindParts (root / "nowhere", true, L".mcam", 0, more).empty (), "a folder that is not there: none");
	fs::remove_all (root, ec);
	}

	// ---- The file name.
	{
	const std::wstring n = Shop::FileName (std::time (nullptr));
	Check (n.rfind (L"Shop_compare_", 0) == 0 && n.size () == std::wstring (L"Shop_compare_20261010.xlsx").size ()
		   && n.find (L"_params") == std::wstring::npos, "Shop_compare_<date>.xlsx - not taken for a part's own dump");
	}

	if (failed)
		{
		std::printf ("shop_test: %d FAILED\n", failed);
		return 1;
		}
	std::puts ("shop_test: all checks passed");
	return 0;
	}
