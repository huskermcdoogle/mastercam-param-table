// The part's .ptconfig: what was typed into a dump's Tools and Summary sheets
// (insert edges, cost, parts per edge, usual edge time; tool edge life and a
// corrected insert name; the batch quantity) and the Program check's ignored
// findings are read back from the saved workbook, kept, and saved.
//
// Usage: partconfig_test <out dir>
//        partconfig_test --harvest <workbook>   (prints what a dump would keep from
//                                                it - tools\check_macros.ps1 uses it)
#include "../src/Csv.h"
#include "../src/PartConfig.h"
#include "../src/Summary.h"
#include "../src/Xlsx.h"

#include <cstdio>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <map>
#include <vector>

static int failed = 0;

static std::string Narrow (const std::wstring &w)
	{
	return Csv::ToUtf8Bom (w).substr (3);
	}

/// Add a sheet to a written workbook (as Excel would save one a macro made): its
/// part, its place in the workbook and its relationship.
static bool AddSheet (const std::filesystem::path &file, const std::string &name, const std::string &sheetXml)
	{
	std::ifstream in (file, std::ios::binary);
	const std::string bytes ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
	in.close ();
	std::map<std::string, std::string> parts;
	std::wstring why;
	if (!Xlsx::Unzip (bytes, parts, why))
		return false;
	auto insertBefore = [] (std::string &s, const std::string &at, const std::string &what)
		{
		const size_t p = s.find (at);
		if (p != std::string::npos)
			s.insert (p, what);
		return p != std::string::npos;
		};
	const bool ok = insertBefore (parts["xl/workbook.xml"], "</sheets>", "<sheet name=\"" + name + "\" sheetId=\"9\" r:id=\"rId99\"/>")
					&& insertBefore (parts["xl/_rels/workbook.xml.rels"], "</Relationships>",
									 "<Relationship Id=\"rId99\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/"
									 "relationships/worksheet\" Target=\"worksheets/sheet9.xml\"/>")
					&& insertBefore (parts["[Content_Types].xml"], "</Types>",
									 "<Override PartName=\"/xl/worksheets/sheet9.xml\" ContentType=\"application/"
									 "vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
	if (!ok)
		return false;
	parts["xl/worksheets/sheet9.xml"] = sheetXml;
	std::vector<std::pair<std::string, std::string>> list (parts.begin (), parts.end ());
	const std::string zip = Xlsx::Zip (list);
	std::ofstream out (file, std::ios::binary | std::ios::trunc);
	out.write (zip.data (), static_cast<std::streamsize> (zip.size ()));
	return static_cast<bool> (out);
	}

/// An inline-string cell.
static std::string Str (const char *ref, const char *text)
	{
	return std::string ("<c r=\"") + ref + "\" t=\"inlineStr\"><is><t>" + text + "</t></is></c>";
	}

/// What a dump would keep from a saved workbook - for tools\check_macros.ps1.
static int HarvestOnly (const char *workbook)
	{
	PartConfig::Config c;
	std::wstring why;
	if (!PartConfig::Harvest (workbook, c, why))
		{
		std::printf ("harvest failed: %s\n", Narrow (why).c_str ());
		return 1;
		}
	for (const auto &kv : c.ignore)
		std::printf ("ignore: %s\n", Narrow (kv.first).c_str ());
	for (const auto &kv : c.inserts)
		if (!kv.second.usualEdgeTime.empty ())
			std::printf ("usual: %s = %s\n", Narrow (kv.first).c_str (), Narrow (kv.second.usualEdgeTime).c_str ());
	return 0;
	}
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
	if (argc > 2 && std::strcmp (argv[1], "--harvest") == 0)
		return HarvestOnly (argv[2]);
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
	s.trackChanges = true;					// a real dump always has its hidden Dumped sheet
	s.untracked.assign (5, 0);

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
	// Inserts table: name, used by, edges, flips, inserts, cost, cost / part, parts per
	// edge, usual edge time.
	s.toolsAfter.push_back ({ Cell (L"Inserts", true), Cell (L"Insert", true), Cell (L"Used by", true),
							  Cell (L"Edges per insert", true), Cell (L"Flips / part", true), Cell (L"Inserts / part", true),
							  Cell (L"Cost per insert", true), Cell (L"Insert cost / part", true), Cell (L"Parts per edge", true),
							  Cell (L"Usual edge time", true) });
	s.toolsAfter.push_back ({ Cell (L""), Cell (L"CNMG 432 typed"), Cell (L"T1"), Cell (L"4"), Cell (L"2"), Cell (L"1"),
							  Cell (L"12.5"), Cell (L""), Cell (L""), Cell (L"12:00") });
	s.toolsAfter.push_back ({ Cell (L""), Cell (L"C 80° r0.031"), Cell (L"T3"), Cell (L"2"), Cell (L"0.2"), Cell (L"0.1"),
							  Cell (L""), Cell (L""), Cell (L"5"), Cell (L"") });
	// The Program check's ignored findings, as a dump writes them from the .ptconfig.
	s.ignored = { { L"no-comment|op 5", L"op 5 has no comment." }, { L"air|op 2", L"op 2 cuts air 40%." } };
	// What the inserts' edges accept (rows 6-7 on the Tools page).
	{
	Xlsx::Sheet::Validation v;
	v.cells = "D6:D7";
	v.type = "whole";
	v.op = "greaterThanOrEqual";
	v.f1 = L"1";
	v.title = L"Edges per insert";
	v.prompt = L"How many edges";
	v.error = L"A whole number, 1 or more.";
	s.toolsValidations.push_back (v);
	}
	Summary::Where w;
	w.title = L"Summary - SAMPLE.mcam";
	w.batchQty = L"25";
	Summary::Add (s, w);
	const std::filesystem::path wb = dir / L"SAMPLE_lathe_params_20261009-120000.xlsx";
	Check (Xlsx::Write (wb.string (), s), "sample workbook written");
	{
	std::map<std::string, std::string> parts;
	std::wstring why;
	Xlsx::Unzip (Xlsx::Build (s), parts, why);
	Check (parts["xl/workbook.xml"].find ("<sheet name=\"Ignored findings\" sheetId=\"5\" state=\"hidden\"") != std::string::npos
			   && parts.count ("xl/worksheets/sheet5.xml"), "the ignored findings go in a hidden sheet");
	}
	// The Program check sheet the macros made, as Excel saves it: its Ignore column
	// set since the last check - air un-ignored, css-no-max ignored.
	Check (AddSheet (wb, "Program check",
					 "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet "
					 "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
					 "<row r=\"1\">" + Str ("A1", "Program check") + "</row>"
					 "<row r=\"2\">" + Str ("A2", "3 findings - 2 ignored (shown greyed at the bottom)") + "</row>"
					 "<row r=\"5\">" + Str ("A5", "#") + Str ("E5", "What was found") + Str ("H5", "Ignore") + Str ("I5", "key") + "</row>"
					 "<row r=\"6\">" + Str ("E6", "op 9 runs CSS with no max spindle speed.") + Str ("H6", "yes")
					 + Str ("I6", "css-no-max|op 9") + "</row>"
					 "<row r=\"7\">" + Str ("E7", "op 2 cuts air 40%.") + Str ("I7", "air|op 2") + "</row>"
					 "<row r=\"8\">" + Str ("E8", "op 5 has no comment.") + Str ("H8", "yes") + Str ("I8", "no-comment|op 5") + "</row>"
					 "</sheetData></worksheet>"),
		   "a Program check sheet added to the workbook");

	// The last dump recorded what it labelled each tool.
	PartConfig::Config c;
	c.tools[L"1"].detected = L"CNMG 432";
	c.tools[L"3"].detected = L"C 80° r0.031";
	c.inserts[L"C 80° r0.031"].cost = L"9.75";		// kept, then cleared in this workbook
	c.inserts[L"C 80° r0.031"].edgesDetected = L"2";	// the dump filled in 2: not a typed value
	c.tools[L"1"].lifeDetected = L"8:00";
	c.tools[L"3"].lifeDetected = L"-";
	std::wstring why;
	Check (PartConfig::Harvest (wb, c, why), "harvest reads the workbook");
	Check (c.tools[L"1"].insert == L"CNMG 432 typed", "a name typed over the label is kept as the tool's insert");
	Check (c.tools[L"3"].insert.empty (), "the label left as dumped is not a correction");
	Check (c.tools[L"1"].edgeLife == L"6:30", "edge life typed for T1");
	Check (c.inserts[L"CNMG 432 typed"].edges == L"4" && c.inserts[L"CNMG 432 typed"].cost == L"12.5", "edges and cost");
	Check (c.inserts[L"C 80° r0.031"].partsPerEdge == L"5", "parts per edge");
	Check (c.inserts[L"C 80° r0.031"].cost.empty (), "a cost cleared in the workbook is cleared");
	Check (c.inserts[L"C 80° r0.031"].edges.empty (), "edges left as the dump filled them in are not kept");
	Check (c.tools[L"3"].edgeLife.empty (), "a blank life the dump left blank is not kept");
	Check (c.batchQty == L"25", "batch quantity from the Summary");
	Check (c.inserts[L"CNMG 432 typed"].usualEdgeTime == L"12:00" && c.inserts[L"C 80° r0.031"].usualEdgeTime.empty (),
		   "usual edge time as typed (blank = none)");
	Check (c.ignore.size () == 2 && c.ignore.count (L"no-comment|op 5") && c.ignore.count (L"css-no-max|op 9")
			   && c.ignore[L"css-no-max|op 9"] == L"op 9 runs CSS with no max spindle speed.",
		   "ignored: the hidden list, then the Program check's Ignore column over it");
	Check (!c.ignore.count (L"air|op 2"), "a finding set back to blank on the Program check is no longer ignored");

	// Without a record of the dump's labels, names are left alone.
	PartConfig::Config fresh;
	PartConfig::Harvest (wb, fresh, why);
	Check (fresh.tools[L"1"].insert.empty (), "no record of labels: names not taken as corrections");

	const std::filesystem::path f = dir / L"SAMPLE.ptconfig";
	c.ignore[L"edge-time|T5|RPGV 1204"] = L"T5 cuts about 24:00 on each edge\r\n(two lines)";
	c.ignore[L"bad=key"] = L"cannot be read back";
	Check (PartConfig::Save (f, c), "saved");
	const PartConfig::Config back = PartConfig::Load (f);
	Check (back.ignore.size () == 3 && back.ignore.count (L"edge-time|T5|RPGV 1204") && back.ignore.count (L"no-comment|op 5")
			   && back.ignore.at (L"edge-time|T5|RPGV 1204") == L"T5 cuts about 24:00 on each edge  (two lines)"
			   && back.inserts.at (L"CNMG 432 typed").usualEdgeTime == L"12:00",
		   "ignored findings (keys as written, a note on one line) and the usual edge time read back; a key with = left out");
	// A workbook with neither the hidden list nor a Program check leaves the kept list alone.
	{
	Xlsx::Sheet bare = s;
	bare.ignored.clear ();
	const std::filesystem::path wb2 = dir / L"BARE_lathe_params.xlsx";
	Xlsx::Write (wb2.string (), bare);
	PartConfig::Config kept;
	kept.ignore[L"needs-regen|op 3"] = L"op 3's toolpath is out of date.";
	PartConfig::Harvest (wb2, kept, why);
	Check (kept.ignore.size () == 1 && kept.ignore.count (L"needs-regen|op 3"), "no list in the workbook: the kept ignores stay");
	}
	Check (back.batchQty == L"25" && back.tools.at (L"1").insert == L"CNMG 432 typed"
			   && back.tools.at (L"1").edgeLife == L"6:30" && back.inserts.at (L"C 80° r0.031").partsPerEdge == L"5"
			   && back.inserts.at (L"C 80° r0.031").edgesDetected == L"2" && back.tools.at (L"3").detected == L"C 80° r0.031"
			   && back.tools.at (L"1").lifeDetected == L"8:00",
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
