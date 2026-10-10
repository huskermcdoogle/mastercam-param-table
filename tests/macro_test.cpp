// Writes macro_sample.xlsm - a small dumped-style sheet carrying the compiled
// macros (res\vbaProject.bin) and the ribbon (vba\ribbon.xml) - and ci_sample.xlsm,
// a program laid out as a dump lays it out (Tools page, inserts table, Summary,
// an ignored finding) for the continuous-improvement tools - for
// tools\check_macros.ps1 to drive in real Excel.
//
// Usage: macro_test <out dir> <vbaProject.bin> <ribbon.xml>
#include "../src/Summary.h"
#include "../src/Xlsx.h"

#include <cstdio>
#include <fstream>
#include <iterator>

namespace
	{
	/// Eight ops in Operation Manager order, six tools, made so that every rule of the
	/// Program check finds exactly one thing (tools\check_macros.ps1 knows the list):
	///   op 3 needs regenerating, has no comment (ignored, from the "part's .ptconfig")
	///   and cuts with the coolant Off; op 5 feeds 0.5 per MIN; op 2 runs CSS with
	///   max_ss 0; op 4 cuts air 40% of 800 s; T1's ROUGH ops run 600 and 500 SFM;
	///   T1 is put in three times; T5 (RPGV 1204) cuts 24:00 an edge, T6 (the same
	///   insert) 10:00. est_seconds: 1:13:00 in all; ops 7, 4, 8 and 1 make 80% of it.
	bool WriteCiSample (const std::string &dir, const std::string &vba, const std::string &ribbon)
		{
		Xlsx::Sheet s;
		//            A          B       C       D           E          F               G             H
		s.rows = { { L"op_idn", L"type", L"tool", L"comment", L"changes", L"needs_regen", L"group_name", L"units",
					 //  I        J            K         L             M          N           O             P
					 L"feed", L"feed_mode", L"speed", L"speed_mode", L"max_ss", L"coolant", L"insp_time", L"flip_longest",
					 //  Q             R           S                  T
					 L"flips_part", L"air_pct", L"cut_seconds_est", L"est_seconds" },
				   { L"1", L"ROUGH", L"1", L"Rough OD", L"0", L"no", L"Main", L"in", L"0.012", L"per rev", L"600", L"CSS", L"3000",
					 L"Flood", L"4:30", L"4:40", L"2", L"10", L"540", L"600" },
				   { L"2", L"FINISH", L"2", L"Finish OD", L"0", L"no", L"Main", L"in", L"0.006", L"per rev", L"800", L"CSS", L"0",
					 L"Flood", L"", L"", L"0", L"", L"240", L"300" },
				   { L"3", L"GROOVE", L"3", L"", L"0", L"yes", L"Main", L"in", L"0.003", L"per rev", L"300", L"RPM", L"2000",
					 L"Off", L"", L"", L"0", L"", L"100", L"120" },
				   { L"4", L"ROUGH", L"1", L"Rough face", L"0", L"no", L"Main", L"in", L"0.012", L"per rev", L"500", L"CSS", L"3000",
					 L"Flood", L"9:00", L"9:00", L"1", L"40", L"800", L"900" },
				   { L"5", L"FINISH", L"4", L"Finish ID", L"0", L"no", L"Main", L"in", L"0.5", L"per min", L"400", L"CSS", L"2500",
					 L"Flood", L"", L"", L"0", L"", L"150", L"200" },
				   { L"6", L"FACE", L"1", L"Face", L"0", L"no", L"Main", L"in", L"0.01", L"per rev", L"600", L"CSS", L"3000",
					 L"Flood", L"", L"", L"0", L"", L"80", L"100" },
				   { L"7", L"ROUGH", L"5", L"Rough OD 2", L"0", L"no", L"Main", L"in", L"0.015", L"per rev", L"700", L"CSS", L"3000",
					 L"Flood", L"24:00", L"24:00", L"1", L"5", L"1440", L"1500" },
				   { L"8", L"ROUGH", L"6", L"Rough ID", L"0", L"no", L"Main", L"in", L"0.015", L"per rev", L"700", L"CSS", L"3000",
					 L"Flood", L"10:00", L"10:00", L"1", L"", L"600", L"660" } };
		const size_t nCols = s.rows[0].size ();
		s.group.assign (nCols, 0);
		s.groupNames = { L"Identity" };
		s.readOnly.assign (nCols, 0);
		s.text = { 0, 1, 0, 1, 0, 1, 1, 1, 0, 1, 0, 1, 0, 1, 1, 1, 0, 0, 0, 0 };
		s.textFormat.assign (nCols, 0);
		s.textFormat[14] = 1;				// insp_time: 9:00 stays 9:00
		s.formula.assign (s.rows.size () - 1, std::vector<std::wstring> (nCols));
		s.frozenCols = 5;
		s.trackChanges = true;
		s.changesCol = 4;
		s.untracked.assign (nCols, 0);
		for (int c : { 16, 18, 19 })		// flips_part, cut_seconds_est, est_seconds: live in a dump
			s.untracked[static_cast<size_t> (c)] = 1;

		// The Tools page as the dump lays it out: D insert, E flips (live), F cut time
		// (live), G longest between flips, H inspection, I edge life; then the inserts
		// table - B insert, C used by, D edges, E flips, F inserts, G cost, H cost / part,
		// I parts per edge, J usual edge time - with the dump's own formulas.
		const wchar_t *numbers[] = { L"1", L"2", L"3", L"4", L"5", L"6" };
		const wchar_t *names[] = { L"OD ROUGH", L"OD FINISH", L"GROOVE", L"ID FINISH", L"OD ROUGH 2", L"ID ROUGH" };
		const wchar_t *inserts[] = { L"CNMG 432", L"VNMG 331", L"Groove 3mm", L"VNMG 331", L"RPGV 1204", L"RPGV 1204" };
		const wchar_t *longest[] = { L"9:00", L"", L"", L"", L"24:00", L"10:00" };
		for (size_t k = 0; k < 6; ++k)
			{
			Xlsx::Sheet::ToolRow t;
			t.number = numbers[k];
			t.name = names[k];
			const std::wstring row = std::to_wstring (k + 2);
			Xlsx::Sheet::FreeCell ins, flips, cut, lng, insp, life;
			ins.text = inserts[k];
			ins.editable = true;
			flips.formula = L"SUMIF('Lathe params'!$C$3:$C$10,$A" + row + L",'Lathe params'!$Q$3:$Q$10)";
			cut.formula = L"TEXT(SUMIF('Lathe params'!$C$3:$C$10,$A" + row + L",'Lathe params'!$S$3:$S$10)/86400,\"[h]:mm:ss\")";
			lng.text = longest[k];
			life.editable = true;
			life.textFormat = true;
			t.extra = { ins, flips, cut, lng, insp, life };
			s.tools.push_back (t);
			}
		s.toolExtraHeads = { L"Insert", L"Flips / part", L"Cut time / part", L"Longest between flips", L"Inspection",
							 L"Edge life (fallback)" };
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
		heads.push_back (head (L"Parts per edge"));
		heads.push_back (head (L"Usual edge time"));
		s.toolsAfter.push_back (heads);
		const wchar_t *tableNames[] = { L"CNMG 432", L"VNMG 331", L"Groove 3mm", L"RPGV 1204", L"" };
		const wchar_t *usedBy[] = { L"T1", L"T2, T4", L"T3", L"T5, T6", L"" };
		const wchar_t *edges[] = { L"4", L"2", L"2", L"8", L"" };
		const wchar_t *costs[] = { L"12.5", L"8", L"15", L"20", L"" };
		for (size_t i = 0; i < 5; ++i)
			{
			const std::wstring r = std::to_wstring (10 + i);		// 6 tools, a blank row, the heading on row 9
			const std::wstring counted = L"SUMIF($D$2:$D$7,$B" + r + L",$E$2:$E$7)";
			Xlsx::Sheet::FreeCell blank, name, used, e, flips, ins, parts, usual;
			name.text = tableNames[i];
			name.editable = true;
			used.text = usedBy[i];
			e.text = edges[i];
			e.editable = true;
			flips.formula = L"IF(AND(N($I" + r + L")>0," + counted + L"<1),1/$I" + r + L"," + counted + L")";
			ins.formula = L"IF(N($D" + r + L")>0,IF(N($E" + r + L")<1,ROUND(N($E" + r + L")/$D" + r + L",4),ROUNDUP(ROUND($E" + r
						  + L"/$D" + r + L",6),0)),\"\")";
			std::vector<Xlsx::Sheet::FreeCell> row = { blank, name, used, e, flips, ins };
			for (Xlsx::Sheet::FreeCell c : Summary::CostCells (10 + i))
				{
				if (c.editable)
					c.text = costs[i];
				row.push_back (c);
				}
			parts.editable = true;
			usual.editable = true;
			usual.textFormat = true;
			row.push_back (parts);
			row.push_back (usual);
			s.toolsAfter.push_back (row);
			}

		// A finding ignored at an earlier check - what the part's .ptconfig hands a dump.
		s.ignored = { { L"no-comment|op 3", L"op 3 has no comment." } };

		Summary::Where w;
		w.title = L"Summary - CIPART.mcam";
		w.subtitle = L"Dumped 2026-10-09 12:00 - 8 operations";
		w.batchQty = L"10";
		Summary::Add (s, w);

		s.vbaProject = vba;
		s.ribbonXml = ribbon;
		return Xlsx::Write (dir + "/ci_sample.xlsm", s);
		}
	}

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
	// X on: the op windows' columns. Ops 2 and 7 are on X-style coolant machines (before /
	// with / after), op 9 on a V9 machine with leftover X-style entries, op 11 on a V9 one;
	// ops 7 and 11 are manual entries; ops 2, 9 and 11 have an inspection comment.
	//            A          B       C       D          E          F              G       H            I        J             K         L           M                   N
	s.rows = { { L"op_idn", L"type", L"tool", L"comment", L"changes", L"tool_radius", L"feed", L"feed_mode", L"speed", L"speed_mode", L"max_ss", L"stepover", L"stepover_percent", L"units", L"coolant_with", L"manual_text",
				 //  Q             R                  S               T                   U              V            W
				 L"insp_time", L"cycle_time_raw", L"est_seconds", L"cut_seconds_est", L"flips_part", L"removed", L"mrr_avg",
				 //  X                  Y                 Z            AA               AB                  AC
				 L"coolant_before", L"coolant_after", L"coolant", L"insp_comment", L"insp_comment_on", L"manual_gcode" },
			   { L"2", L"DYNAMIC", L"1", L"Rough OD", L"0", L"0.5", L"0.01", L"per rev", L"200", L"CSS", L"3500", L"0.25", L"50", L"in", L"none", L"",
				 L"8:00", L"660", L"660", L"600", L"1", L"2", L"0.182",
				 L"none", L"none", L"", L"ROTATE INSERT", L"1", L"" },
			   { L"7", L"FINISH", L"3", L"Finish OD", L"0", L"", L"0.008", L"per rev", L"300", L"CSS", L"3500", L"", L"", L"in", L"none", L"G4 X1.\r\nM01",
				 L"", L"330", L"330", L"300", L"0", L"0.5", L"0.091",
				 L"none", L"none", L"", L"", L"", L"1006" },
			   { L"9", L"ROUGH", L"1", L"Rough face", L"0", L"", L"0.012", L"per rev", L"250", L"CSS", L"3000", L"", L"", L"in", L"Flood", L"",
				 L"8:00", L"990", L"990", L"900", L"1", L"3", L"0.182",
				 L"none", L"none", L"Flood", L"CHECK INSERT", L"1", L"" },
			   { L"11", L"FINISH", L"1", L"Finish face", L"0", L"", L"0.006", L"per rev", L"350", L"CSS", L"3000", L"", L"", L"in", L"", L"(TOOL CHECK)",
				 L"", L"260", L"260", L"240", L"0", L"0.2", L"0.046",
				 L"", L"", L"Off", L"", L"0", L"1005" } };
	s.group = { 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 2, 2, 0, 3, 3, 3, 4, 4, 4, 4, 4, 4, 5, 5, 5, 6, 6, 6 };
	s.groupNames = { L"Identity", L"Feeds", L"Depth", L"Other", L"Stats", L"Coolant", L"Inspection" };
	s.readOnly = { 1, 1, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0 };
	s.text = { 0, 1, 0, 1, 0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0 };
	s.notApplicable = { { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1 },
						{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0 },
						{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1 },
						{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0 } };
	s.formula.assign (4, std::vector<std::wstring> (29));
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
	s.outlineGroup = { 0, 1, 1, 0, 0, 0, 0 };	// Feeds and Depth fold (Go to)
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
	rule ("A3:C6 E3:F6 N3:N6 R3:W6", "custom", "", L"FALSE", L"");	// read-only
	rule ("L4:M6 P3 P5 O6 X6:Y6 Z3:Z4 AA4:AB4 AC3 AC5", "custom", "", L"FALSE", L"");	// does not apply
	rule ("P4 P6", "textLength", "lessThanOrEqual", L"3111", L"");
	rule ("D3:D6", "textLength", "lessThanOrEqual", L"119", L"");		// the op comment
	rule ("AA3 AA5:AA6", "textLength", "lessThanOrEqual", L"49", L"");	// the inspection stop's
	rule ("AC4 AC6", "whole", "between", L"1005", L"1006");
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
	// Two X-style machines: one has Thru-tool, one does not. Op 9's V9 machine still
	// carries X-style entries - all it may do with them is clear them (none).
	v.cells = "O3 X3 Y3";
	v.choices = { L"none", L"Flood", L"Mist", L"Thru-tool" };
	s.validations.push_back (v);
	v.cells = "O4 X4 Y4";
	v.choices = { L"none", L"Flood", L"Mist" };
	s.validations.push_back (v);
	v.cells = "O5 X5 Y5";
	v.choices = { L"none" };
	s.validations.push_back (v);
	v.cells = "Z5:Z6";
	v.choices = { L"Off", L"Flood", L"Mist", L"Thru-tool" };
	s.validations.push_back (v);
	v.cells = "AB3 AB5:AB6";
	v.choices = { L"0", L"1" };
	s.validations.push_back (v);
	}

	// A Tools page with live columns and an inserts table under it, and a column
	// formatted as Text (what keeps a typed 9:00 from becoming a time of day).
	// The percent is a calculated column that follows the stepover: untracked, so
	// it is not an edit - drawn blue when it has moved.
	s.untracked.assign (29, 0);
	s.untracked[12] = 1;
	for (int c : { 18, 19, 20, 22 })
		s.untracked[static_cast<size_t> (c)] = 1;
	s.textFormat.assign (29, 0);
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
	Xlsx::Sheet::FreeCell ins, flips, ic, angle;
	ins.text = L"CNMG 432";
	flips.formula = L"SUMIF('Lathe params'!$C$3:$C$4,$A2,'Lathe params'!$I$3:$I$4)";
	flips.text = L"200";
	// The chip thinning's insert size and entering angle (found by heading): T1 no
	// angle - its DYNAMIC op 2 cuts with a 1.0 round (tool radius 0.5); T3 a
	// 95-degree holder (PCLNR / MCLNR).
	ic.text = L"1";
	t1.extra = { ins, flips, ic, angle };
	flips.formula = L"SUMIF('Lathe params'!$C$3:$C$4,$A3,'Lathe params'!$I$3:$I$4)";
	flips.text = L"300";
	ic.text = L"0.5";
	angle.text = L"95";
	t3.extra = { ins, flips, ic, angle };
	s.tools = { t1, t3 };
	s.toolExtraHeads = { L"Insert", L"Flips / part", L"Insert size (IC)", L"Entering angle" };
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
	if (!WriteCiSample (argv[1], s.vbaProject, s.ribbonXml))
		{
		std::puts ("FAIL: write ci_sample.xlsm");
		return 1;
		}
	std::puts ("macro_test: wrote macro_sample.xlsm and ci_sample.xlsm");
	return 0;
	}
