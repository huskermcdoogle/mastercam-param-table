// Writes tests\out\sample.xlsx for tools\check_xlsx.py to verify with openpyxl.
#include "../src/Xlsx.h"

#include <cstdio>

int main (int argc, char **argv)
	{
	Xlsx::Sheet s;
	s.rows = {
		{ L"op_idn", L"type", L"tool", L"comment", L"step", L"direction", L"feed", L"tplane_id" },
		{ L"1", L"ROUGH", L"1", L"R & <T> \"q\" é", L"0.1", L"0", L"-0.005", L"0" },
		{ L"2", L"FINISH", L"2", L"", L"", L"1", L"12.5", L"0" },
		{ L"3", L"DYNAMIC", L"3", L"007", L"1e-3", L"", L"100", L"1" },
	};
	s.group = { 0, 0, 0, 0, 1, 1, 2, 3 };
	s.groupNames = { L"identity", L"toolpath", L"feeds", L"where" };
	s.readOnly = { 1, 1, 0, 0, 0, 1, 0, 1 };
	s.text = { 0, 1, 0, 1, 0, 0, 0, 0 };
	s.notApplicable = { { 0, 0, 0, 0, 0, 0, 0, 0 }, { 0, 0, 0, 0, 1, 0, 0, 0 },
						{ 0, 0, 0, 0, 0, 1, 0, 0 } };
	s.formula.assign (3, std::vector<std::wstring> (8));
	s.formula[0][4] = L"IF(G3=-0.005,0.1,G3/$C3*100)";
	{
	Xlsx::Sheet::Validation v;
	v.cells = "F3 F5";
	v.type = "list";
	v.choices = { L"none", L"10BAR", L"70 bar (M216)" };
	v.title = L"coolant_with";
	v.prompt = L"Coolant on WITH the move.\nPick from the list.";
	s.validations.push_back (v);
	Xlsx::Sheet::Validation n;
	n.cells = "G3:G5";
	n.type = "decimal";
	n.op = "greaterThanOrEqual";
	n.f1 = L"0";
	n.title = L"feed";
	n.prompt = L"A number, 0 or more.";
	n.error = L"A number, 0 or more.";
	s.validations.push_back (n);
	}
	{
	// A 2x2 red PNG.
	static const unsigned char png[] = {
		0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
		0x00,0x00,0x00,0x02,0x00,0x00,0x00,0x02,0x08,0x02,0x00,0x00,0x00,0xFD,0xD4,0x9A,
		0x73,0x00,0x00,0x00,0x10,0x49,0x44,0x41,0x54,0x78,0x9C,0x63,0xB8,0xE0,0x60,0x00,
		0x44,0x0C,0x10,0x0A,0x00,0x25,0x8E,0x05,0x01,0x98,0x1A,0xB1,0xEA,0x00,0x00,0x00,
		0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82 };
	Xlsx::Sheet::ToolRow t;
	t.number = L"1";
	t.name = L"OD rough 80 deg";
	t.usedBy = L"op 1";
	t.png.assign (reinterpret_cast<const char *> (png), sizeof (png));
	t.width = 2;
	t.height = 2;
	s.tools.push_back (t);
	Xlsx::Sheet::ToolRow u;
	u.number = L"7";
	u.name = L"mill tool - no picture";
	s.tools.push_back (u);
	s.toolLinks.push_back ({ "C3", 0 });
	}
	s.trackChanges = true;
	s.changesCol = 7;			// a column no formula refers to (C is the percent formula's radius)
	s.outlineGroup = { 0, 1, 1, 0 };
	s.collapseGroup = { 0, 0, 1, 0 };
	s.frozenCols = 4;
	const char *out = argc > 1 ? argv[1] : "sample.xlsx";
	if (!Xlsx::Write (out, s))
		{
		std::puts ("FAIL: write");
		return 1;
		}
	if (Xlsx::ColName (0) != "A" || Xlsx::ColName (26) != "AA" || Xlsx::ColName (25) != "Z"
		|| Xlsx::ColName (701) != "ZZ" || Xlsx::ColName (702) != "AAA")
		{
		std::puts ("FAIL: ColName");
		return 1;
		}
	std::puts ("xlsx_test ok");
	return 0;
	}
