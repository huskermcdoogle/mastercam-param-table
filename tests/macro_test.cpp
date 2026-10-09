// Writes macro_sample.xlsm - a small dumped-style sheet carrying the compiled
// macros (res\vbaProject.bin) and the ribbon (vba\ribbon.xml) - for
// tools\check_macros.ps1 to drive in real Excel.
//
// Usage: macro_test <out dir> <vbaProject.bin> <ribbon.xml>
#include "../src/Summary.h"
#include "../src/Xlsx.h"

#include <cstdio>
#include <fstream>
#include <iterator>

int main (int argc, char **argv)
	{
	if (argc < 4)
		{
		std::puts ("usage: macro_test <out dir> <vbaProject.bin> <ribbon.xml>");
		return 1;
		}
	auto slurp = [] (const char *p)
		{
		std::ifstream in (p, std::ios::binary);
		return std::string ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		};

	Xlsx::Sheet s;
	//            A          B       C       D          E          F              G       H            I        J             K         L           M                   N
	s.rows = { { L"op_idn", L"type", L"tool", L"comment", L"changes", L"tool_radius", L"feed", L"feed_mode", L"speed", L"speed_mode", L"max_ss", L"stepover", L"stepover_percent", L"units", L"coolant_with", L"manual_text" },
			   { L"2", L"DYNAMIC", L"1", L"Rough OD", L"0", L"0.5", L"0.01", L"per rev", L"200", L"CSS", L"3500", L"0.25", L"50", L"in", L"none", L"" },
			   { L"7", L"FINISH", L"3", L"Finish OD", L"0", L"", L"0.008", L"per rev", L"300", L"CSS", L"3500", L"", L"", L"in", L"none", L"G4 X1.\r\nM01" } };
	s.group = { 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 0, 3, 3 };
	s.groupNames = { L"Identity", L"Feeds", L"Depth", L"Other" };
	s.readOnly = { 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0 };
	s.text = { 0, 1, 0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, 1 };
	s.notApplicable = { { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1 },
						{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0 } };
	s.formula.assign (2, std::vector<std::wstring> (16));
	s.formula[0][12] = L"IF(L3=0.25,50,L3/$F3*100)";
	s.frozenCols = 5;
	s.trackChanges = true;
	s.changesCol = 4;

	auto rule = [&s] (const char *cells, const char *type, const char *op, const wchar_t *f1, const wchar_t *f2)
		{
		Xlsx::Sheet::Validation v;
		v.cells = cells;
		v.type = type;
		v.op = op;
		v.f1 = f1;
		v.f2 = f2;
		s.validations.push_back (v);
		};
	rule ("A3:F4 N3:N4", "custom", "", L"FALSE", L"");				// read-only
	rule ("L4:M4 P3", "custom", "", L"FALSE", L"");				// does not apply
	rule ("P4", "textLength", "lessThanOrEqual", L"3111", L"");
	rule ("G3:G4 L3 M3", "decimal", "greaterThanOrEqual", L"0", L"");
	rule ("I3:I4 K3:K4", "whole", "greaterThanOrEqual", L"0", L"");
	{
	Xlsx::Sheet::Validation v;
	v.cells = "H3:H4";
	v.type = "list";
	v.choices = { L"per rev", L"per min" };
	s.validations.push_back (v);
	v.cells = "J3:J4";
	v.choices = { L"RPM", L"CSS" };
	s.validations.push_back (v);
	// Two machines: one has Thru-tool, one does not.
	v.cells = "O3";
	v.choices = { L"none", L"Flood", L"Mist", L"Thru-tool" };
	s.validations.push_back (v);
	v.cells = "O4";
	v.choices = { L"none", L"Flood", L"Mist" };
	s.validations.push_back (v);
	}

	// A Tools page with live columns and an inserts table under it, and a column
	// formatted as Text (what keeps a typed 9:00 from becoming a time of day).
	// The percent is a calculated column that follows the stepover: untracked, so
	// it is not an edit - drawn blue when it has moved.
	s.untracked.assign (16, 0);
	s.untracked[12] = 1;
	s.textFormat.assign (16, 0);
	s.textFormat[15] = 1;
	{
	Xlsx::Sheet::ToolRow t1, t3;
	t1.number = L"1";
	t1.name = L"OD ROUGH";
	t1.usedBy = L"op 2";
	t3.number = L"3";
	t3.name = L"OD FINISH";
	t3.usedBy = L"op 7";
	Xlsx::Sheet::FreeCell ins, flips;
	ins.text = L"CNMG 432";
	flips.formula = L"SUMIF('Lathe params'!$C$3:$C$4,$A2,'Lathe params'!$I$3:$I$4)";
	flips.text = L"200";
	t1.extra = { ins, flips };
	flips.formula = L"SUMIF('Lathe params'!$C$3:$C$4,$A3,'Lathe params'!$I$3:$I$4)";
	flips.text = L"300";
	t3.extra = { ins, flips };
	s.tools = { t1, t3 };
	s.toolExtraHeads = { L"Insert", L"Flips / part" };
	Xlsx::Sheet::FreeCell h, name, edges, total, per;
	h.head = true;
	h.text = L"Insert";
	name.text = L"CNMG 432";
	edges.text = L"4";
	edges.editable = true;
	total.formula = L"SUMIF($D$2:$D$3,$B6,$E$2:$E$3)";
	total.text = L"500";
	per.formula = L"IF(N($C6)>0,ROUND($D6/$C6,2),\"\")";
	per.text = L"125";
	s.toolsAfter = { { Xlsx::Sheet::FreeCell (), h }, { Xlsx::Sheet::FreeCell (), name, edges, total, per } };
	}

	// The Summary in front, as a dump writes it: the macros must still find
	// their sheets (by name, and the main sheet by its code name Sheet1).
	Summary::Where w;
	w.title = L"Summary - macro sample";
	Summary::Add (s, w);

	s.vbaProject = slurp (argv[2]);
	s.ribbonXml = slurp (argv[3]);
	if (s.vbaProject.empty () || s.ribbonXml.empty ())
		{
		std::puts ("FAIL: no compiled macros or ribbon to embed");
		return 1;
		}
	const std::string out = std::string (argv[1]) + "/macro_sample.xlsm";
	if (!Xlsx::Write (out, s))
		{
		std::puts ("FAIL: write");
		return 1;
		}
	std::puts ("macro_test: wrote macro_sample.xlsm");
	return 0;
	}
