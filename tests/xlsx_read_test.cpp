// The xlsx reader against fixtures made by Python's own zlib and zipfile
// (tools/make_xlsx_fixtures.py), and against this tool's own writer.
#include "../src/Xlsx.h"

#include <cstdio>
#include <fstream>
#include <iterator>
#include <string>

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

	std::string Slurp (const std::string &path)
		{
		std::ifstream in (path, std::ios::binary);
		return std::string ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		}
	}

int main (int argc, char **argv)
	{
	const std::string dir = argc > 1 ? argv[1] : "fixtures";

	// ---- DEFLATE: every block type, matched byte for byte against zlib.
	for (int k = 1; k <= 6; ++k)
		{
		const std::string packed = Slurp (dir + "/deflate_" + std::to_string (k) + ".bin");
		const std::string want = Slurp (dir + "/deflate_" + std::to_string (k) + ".txt");
		std::string got;
		const bool ok = Xlsx::Inflate (packed, got);
		char what[64];
		std::snprintf (what, sizeof (what), "inflate sample %d", k);
		Check (ok && got == want, what);
		}
	{
	std::string got;
	Check (!Xlsx::Inflate (std::string ("\x07\xff\xff", 3), got), "damaged deflate data is refused");
	}

	// ---- A workbook laid out the way Excel saves one.
	{
	std::vector<std::vector<std::wstring>> rows;
	std::vector<size_t> sheetRow;
	std::wstring why;
	const bool ok = Xlsx::ReadSheet (dir + "/excel_like.xlsx", rows, sheetRow, why);
	Check (ok, "excel-like workbook reads");
	if (ok)
		{
		Check (rows.size () == 3, "column names + two data rows (gap and blank rows dropped)");
		Check (sheetRow.size () == 3 && sheetRow[0] == 2 && sheetRow[1] == 3 && sheetRow[2] == 5,
			   "each row keeps its Excel row number");
		Check (rows[0].size () == 6 && rows[0][0] == L"op_idn" && rows[0][5] == L"manual_text",
			   "starts at the column-name row, not the group row above it");
		Check (rows[1][1] == L"ROUGH", "a shared string");
		Check (rows[1][2] == L"Datum -A-", "rich text runs joined, phonetic guide dropped");
		Check (rows[1][3] == L"0.01", "a number in its shortest exact form, not as Excel wrote it");
		Check (rows[1][4] == L"per rev", "a dropdown choice");
		Check (rows[1][5] == L"M00\r\nM01", "Excel's _x000D_ escape is a carriage return");
		Check (rows[2][0] == L"7", "a formula reads as its calculated value");
		Check (rows[2][2] == L"R & D <1>", "entities decoded");
		Check (rows[2][3] == L"0.5", "a formula returning text");
		Check (rows[2][4] == L"1", "a boolean");
		Check (rows[2][5] == L"#N/A", "an error comes through as text (and a load refuses it)");
		}
	}

	// ---- This tool's own output reads back to what was written.
	{
	Xlsx::Sheet s;
	s.rows = { { L"op_idn", L"type", L"comment", L"feed" },
			   { L"1", L"ROUGH", L"a < b & \"c\"", L"0.008" },
			   { L"2", L"FINISH", L"", L"12.5" } };
	s.group = { 0, 0, 0, 1 };
	s.groupNames = { L"Identity", L"Feeds" };
	s.formula.assign (2, std::vector<std::wstring> (4));
	s.formula[1][3] = L"IF(D4=12.5,12.5,0)";
	std::vector<std::vector<std::wstring>> rows;
	std::vector<size_t> sheetRow;
	std::wstring why;
	const bool ok = Xlsx::ReadSheetBytes (Xlsx::Build (s), rows, sheetRow, why);
	Check (ok && rows.size () == 3, "own workbook reads back");
	if (ok && rows.size () == 3)
		{
		Check (rows[0] == s.rows[0], "own column names");
		Check (rows[1] == s.rows[1], "own row: inline strings, escapes, numbers");
		Check (rows[2][3] == L"12.5", "own formula cell reads as its cached value");
		Check (sheetRow[1] == 3, "own data starts on Excel row 3");
		}
	}

	{
	std::vector<std::vector<std::wstring>> rows;
	std::vector<size_t> sheetRow;
	std::wstring why;
	Check (!Xlsx::ReadSheetBytes ("not a zip at all", rows, sheetRow, why) && !why.empty (),
		   "a non-workbook is refused with a reason");
	}

	if (gFailed != 0)
		{
		std::printf ("xlsx_read_test: %d check(s) FAILED\n", gFailed);
		return 1;
		}
	std::printf ("xlsx_read_test: all checks passed\n");
	return 0;
	}
