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
	// Four ops, three of them on tool 1. Q on: tool inspection and a live estimate shaped
	// like the dump's - est_seconds = the fixed time + cut_seconds_est, the cut time going
	// inversely with feed and speed; flips_part = cut time / insp_time; mrr_avg = removed /
	// est time - so the planning macros have something real to work on.
	//            A          B       C       D          E          F              G       H            I        J             K         L           M                   N
	s.rows = { { L"op_idn", L"type", L"tool", L"comment", L"changes", L"tool_radius", L"feed", L"feed_mode", L"speed", L"speed_mode", L"max_ss", L"stepover", L"stepover_percent", L"units", L"coolant_with", L"manual_text",
				 //  Q             R                  S               T                   U              V            W
				 L"insp_time", L"cycle_time_raw", L"est_seconds", L"cut_seconds_est", L"flips_part", L"removed", L"mrr_avg" },
			   { L"2", L"DYNAMIC", L"1", L"Rough OD", L"0", L"0.5", L"0.01", L"per rev", L"200", L"CSS", L"3500", L"0.25", L"50", L"in", L"none", L"",
				 L"8:00", L"660", L"660", L"600", L"1", L"2", L"0.182" },
			   { L"7", L"FINISH", L"3", L"Finish OD", L"0", L"", L"0.008", L"per rev", L"300", L"CSS", L"3500", L"", L"", L"in", L"none", L"G4 X1.\r\nM01",
				 L"", L"330", L"330", L"300", L"0", L"0.5", L"0.091" },
			   { L"9", L"ROUGH", L"1", L"Rough face", L"0", L"", L"0.012", L"per rev", L"250", L"CSS", L"3000", L"", L"", L"in", L"none", L"",
				 L"8:00", L"990", L"990", L"900", L"1", L"3", L"0.182" },
			   { L"11", L"FINISH", L"1", L"Finish face", L"0", L"", L"0.006", L"per rev", L"350", L"CSS", L"3000", L"", L"", L"in", L"none", L"",
				 L"", L"260", L"260", L"240", L"0", L"0.2", L"0.046" } };
	s.group = { 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 0, 3, 3, 3, 4, 4, 4, 4, 4, 4 };
	s.groupNames = { L"Identity", L"Feeds", L"Depth", L"Other", L"Stats" };
	s.readOnly = { 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 1, 1, 1 };
	s.text = { 0, 1, 0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0 };
	s.notApplicable = { { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0 },
						{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 },
						{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0 },
						{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0 } };
	s.formula.assign (4, std::vector<std::wstring> (23));
	s.formula[0][12] = L"IF(L3=0.25,50,L3/$F3*100)";
	{
	// The cut time at the dumped feed and speed, and the time feeds do not move.
	const wchar_t *cut0[] = { L"600", L"300", L"900", L"240" }, *feed0[] = { L"0.01", L"0.008", L"0.012", L"0.006" };
	const wchar_t *speed0[] = { L"200", L"300", L"250", L"350" }, *fixed[] = { L"60", L"30", L"90", L"20" };
	for (int d = 0; d < 4; ++d)
		{
		const std::wstring r = std::to_wstring (d + 3), q = L"Q" + r;
		// insp_time as seconds - the dump's own SecondsOf.
		const std::wstring secs = L"IF(ISNUMBER(" + q + L"),IF(" + q + L"<1," + q + L"*86400," + q + L"),IFERROR(IF(LEN(" + q
								  + L")-LEN(SUBSTITUTE(" + q + L",\":\",\"\"))=1,TIMEVALUE(\"0:\"&" + q + L"),TIMEVALUE(" + q
								  + L"))*86400,VALUE(" + q + L")))";
		s.formula[d][19] = L"IFERROR(" + std::wstring (cut0[d]) + L"*" + feed0[d] + L"/G" + r + L"*" + speed0[d] + L"/I" + r + L",\"\")";
		s.formula[d][18] = std::wstring (fixed[d]) + L"+T" + r;
		s.formula[d][20] = L"IF(" + q + L"=\"\",0,IFERROR(INT(T" + r + L"/" + secs + L"),0))";
		s.formula[d][22] = L"IFERROR(ROUND(V" + r + L"/S" + r + L"*60,3),\"\")";
		}
	}
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
	rule ("A3:F6 N3:N6 R3:W6", "custom", "", L"FALSE", L"");		// read-only
	rule ("L4:M6 P3 P5:P6", "custom", "", L"FALSE", L"");			// does not apply
	rule ("P4", "textLength", "lessThanOrEqual", L"3111", L"");
	rule ("G3:G4 G6 L3 M3", "decimal", "greaterThanOrEqual", L"0", L"");
	rule ("G5", "decimal", "between", L"0", L"0.013");				// a feed with a ceiling
	rule ("I3:I6 K3:K6", "whole", "greaterThanOrEqual", L"0", L"");
	{
	Xlsx::Sheet::Validation v;
	v.cells = "H3:H6";
	v.type = "list";
	v.choices = { L"per rev", L"per min" };
	s.validations.push_back (v);
	v.cells = "J3:J6";
	v.choices = { L"RPM", L"CSS" };
	s.validations.push_back (v);
	// Two machines: one has Thru-tool, one does not.
	v.cells = "O3";
	v.choices = { L"none", L"Flood", L"Mist", L"Thru-tool" };
	s.validations.push_back (v);
	v.cells = "O4:O6";
	v.choices = { L"none", L"Flood", L"Mist" };
	s.validations.push_back (v);
	}

	// A Tools page with live columns and an inserts table under it, and a column
	// formatted as Text (what keeps a typed 9:00 from becoming a time of day).
	// The percent is a calculated column that follows the stepover: untracked, so
	// it is not an edit - drawn blue when it has moved.
	s.untracked.assign (23, 0);
	s.untracked[12] = 1;
	for (int c : { 18, 19, 20, 22 })
		s.untracked[static_cast<size_t> (c)] = 1;
	s.textFormat.assign (23, 0);
	s.textFormat[15] = 1;
	s.textFormat[16] = 1;			// insp_time: a typed 8:00 stays 8:00
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
