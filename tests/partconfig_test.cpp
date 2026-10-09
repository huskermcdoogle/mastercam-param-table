// The part's .ptconfig: what was typed into a dump's Tools and Summary sheets
// (insert edges, cost, parts per edge; tool edge life and a corrected insert name;
// the batch quantity) is read back from the saved workbook, kept, and saved.
#include "../src/PartConfig.h"
#include "../src/Summary.h"
#include "../src/Xlsx.h"

#include <cstdio>
#include <filesystem>
#include <fstream>

static int failed = 0;
static void Check (bool ok, const char *what)
	{
	std::printf ("  %s  %s\n", ok ? "ok  " : "FAIL", what);
	if (!ok)
		++failed;
	}

static Xlsx::Sheet::FreeCell Cell (const wchar_t *t, bool head = false)
	{
	Xlsx::Sheet::FreeCell c;
	c.text = t;
	c.head = head;
	c.editable = !head;
	return c;
	}

int main (int argc, char **argv)
	{
	const std::filesystem::path dir = argc > 1 ? argv[1] : ".";
	Xlsx::Sheet s;
	s.rows = { { L"op_idn", L"type", L"tool", L"est_seconds", L"flips_part" },
			   { L"2", L"ROUGH", L"1", L"300", L"2" },
			   { L"5", L"FINISH", L"3", L"120", L"0" } };
	s.group.assign (5, 0);
	s.groupNames = { L"Identity" };
	s.readOnly.assign (5, 1);
	s.text = { 0, 1, 0, 0, 0 };
	s.formula.assign (2, std::vector<std::wstring> (5));

	// Tools: D insert (T1 corrected by hand, T3 as the dump labelled it), H edge life.
	Xlsx::Sheet::ToolRow t1, t3;
	t1.number = L"1";
	t1.name = L"OD ROUGH";
	t3.number = L"3";
	t3.name = L"OD FINISH";
	t1.extra = { Cell (L"CNMG 432 typed"), Cell (L"2"), Cell (L"0:05:00"), Cell (L""), Cell (L"6:30") };
	t3.extra = { Cell (L"C 80° r0.031"), Cell (L"0"), Cell (L"0:02:00"), Cell (L""), Cell (L"") };
	s.tools = { t1, t3 };
	s.toolExtraHeads = { L"Insert", L"Flips / part", L"Cut time / part", L"Longest between flips", L"Edge life (fallback)" };
	// Inserts table: name, used by, edges, flips, inserts, cost, cost / part, parts per edge.
	s.toolsAfter.push_back ({ Cell (L"Inserts", true), Cell (L"Insert", true), Cell (L"Used by", true),
							  Cell (L"Edges per insert", true), Cell (L"Flips / part", true), Cell (L"Inserts / part", true),
							  Cell (L"Cost per insert", true), Cell (L"Insert cost / part", true), Cell (L"Parts per edge", true) });
	s.toolsAfter.push_back ({ Cell (L""), Cell (L"CNMG 432 typed"), Cell (L"T1"), Cell (L"4"), Cell (L"2"), Cell (L"1"),
							  Cell (L"12.5"), Cell (L""), Cell (L"") });
	s.toolsAfter.push_back ({ Cell (L""), Cell (L"C 80° r0.031"), Cell (L"T3"), Cell (L"2"), Cell (L"0.2"), Cell (L"0.1"),
							  Cell (L""), Cell (L""), Cell (L"5") });
	Summary::Where w;
	w.title = L"Summary - SAMPLE.mcam";
	w.batchQty = L"25";
	Summary::Add (s, w);
	const std::filesystem::path wb = dir / L"SAMPLE_lathe_params_20261009-120000.xlsx";
	Check (Xlsx::Write (wb.string (), s), "sample workbook written");

	// The last dump recorded what it labelled each tool.
	PartConfig::Config c;
	c.tools[L"1"].detected = L"CNMG 432";
	c.tools[L"3"].detected = L"C 80° r0.031";
	c.inserts[L"C 80° r0.031"].cost = L"9.75";		// typed in an earlier dump, blank here: kept
	std::wstring why;
	Check (PartConfig::Harvest (wb, c, why), "harvest reads the workbook");
	Check (c.tools[L"1"].insert == L"CNMG 432 typed", "a name typed over the label is kept as the tool's insert");
	Check (c.tools[L"3"].insert.empty (), "the label left as dumped is not a correction");
	Check (c.tools[L"1"].edgeLife == L"6:30", "edge life typed for T1");
	Check (c.inserts[L"CNMG 432 typed"].edges == L"4" && c.inserts[L"CNMG 432 typed"].cost == L"12.5", "edges and cost");
	Check (c.inserts[L"C 80° r0.031"].partsPerEdge == L"5", "parts per edge");
	Check (c.inserts[L"C 80° r0.031"].cost == L"9.75", "a blank cell does not erase what was kept");
	Check (c.batchQty == L"25", "batch quantity from the Summary");

	// Without a record of the dump's labels, names are left alone.
	PartConfig::Config fresh;
	PartConfig::Harvest (wb, fresh, why);
	Check (fresh.tools[L"1"].insert.empty (), "no record of labels: names not taken as corrections");

	const std::filesystem::path f = dir / L"SAMPLE.ptconfig";
	Check (PartConfig::Save (f, c), "saved");
	const PartConfig::Config back = PartConfig::Load (f);
	Check (back.batchQty == L"25" && back.tools.at (L"1").insert == L"CNMG 432 typed"
			   && back.tools.at (L"1").edgeLife == L"6:30" && back.inserts.at (L"C 80° r0.031").partsPerEdge == L"5"
			   && back.inserts.at (L"C 80° r0.031").cost == L"9.75" && back.tools.at (L"3").detected == L"C 80° r0.031",
		   "read back the same (names with spaces and the degree sign)");
	Check (PartConfig::PathFor (L"C:\\parts\\Oil Spool.mcam") == std::filesystem::path (L"C:\\parts\\Oil Spool.ptconfig"),
		   "the file sits beside the part");
	Check (PartConfig::NewestDump (dir / L"SAMPLE.mcam", std::filesystem::path ()) == wb, "the newest dump of the part is found");
	Check (PartConfig::NewestDump (dir / L"OTHER.mcam", std::filesystem::path ()).empty (), "another part's dumps are not");

	if (failed)
		std::printf ("partconfig_test: %d FAILED\n", failed);
	else
		std::printf ("partconfig_test: all checks passed\n");
	return failed ? 1 : 0;
	}
