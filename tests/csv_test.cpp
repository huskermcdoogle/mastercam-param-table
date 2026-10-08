//
// csv_test.cpp - the CSV and cell-value handling, outside Mastercam.
//
#include "../src/Csv.h"

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
	}

int main ()
	{
	using namespace Csv;

	// ---- A round trip keeps every awkward cell exactly.
	{
	std::vector<Row> in = {
		{ L"op", L"description", L"step" },
		{ L"12", L"Rough OD, 80 deg", L"0.1" },
		{ L"13", L"say \"hi\"", L"" },
		{ L"14", L"two\r\nlines", L"  padded  " },
		{ L"15", L"deg \x00B0 sign", L"1e-05" },
	};
	const std::vector<Row> out = Parse (Serialize (in));
	Check (out == in, "serialize then parse returns the same rows");
	}

	// ---- What Excel and Notepad each do to a file.
	{
	const std::vector<Row> a = Parse (L"\xFEFF" L"a,b\r\n1,2\r\n");
	Check (a.size () == 2 && a[0][0] == L"a", "a leading BOM is ignored");

	const std::vector<Row> b = Parse (L"a,b\n1,2");
	Check (b.size () == 2 && b[1][1] == L"2", "bare LF and no final newline");

	const std::vector<Row> c = Parse (L"a,b\r\n\r\n\r\n1,2\r\n");
	Check (c.size () == 2, "blank lines are not rows");

	const std::vector<Row> d = Parse (L"a,,c\r\n");
	Check (d.size () == 1 && d[0].size () == 3 && d[0][1].empty (),
		   "an empty cell in the middle is kept");

	const std::vector<Row> e = Parse (L"a,\"\"\r\n");
	Check (e.size () == 1 && e[0].size () == 2,
		   "a quoted empty cell at the end is kept");
	}

	// ---- UTF-8 with a BOM, and the ANSI fallback.
	{
	const std::wstring text = L"deg \x00B0 and caf\x00E9";
	const std::string bytes = ToUtf8Bom (text);
	Check (bytes.size () >= 3 && (unsigned char) bytes[0] == 0xEF,
		   "output carries a BOM");
	Check (FromUtf8 (bytes) == text, "UTF-8 round trips");

	// Excel saving as plain ANSI: 0xB0 alone is a degree sign there, and is
	// not valid UTF-8.
	Check (FromUtf8 (std::string ("deg \xB0")) == L"deg \x00B0",
		   "ANSI input falls back instead of failing");
	}

	// ---- Doubles: shortest text, exact round trip.
	{
	Check (FormatDouble (0.1) == L"0.1", "0.1 prints as 0.1");
	Check (FormatDouble (0.0) == L"0", "zero prints as 0");
	Check (FormatDouble (-0.0) == L"0", "negative zero prints as 0");
	Check (FormatDouble (2.0) == L"2", "an integral double has no point");
	Check (FormatDouble (0.03125) == L"0.03125", "0.03125 prints in full");

	// Tidy: Mastercam's arithmetic noise dropped, and still the same value.
	Check (Tidy (0.2500000000000021) == L"0.25", "tidy drops noise above");
	Check (Tidy (0.03000000000000037) == L"0.03", "tidy drops noise, small");
	Check (Tidy (12.000000000000147) == L"12", "tidy drops noise to an integer");
	Check (Tidy (-0.33020000000000005) == L"-0.3302", "tidy keeps the sign");
	Check (Tidy (0.00001) == L"0.00001", "tidy never uses an exponent");
	Check (Tidy (0.0) == L"0", "tidy zero");
	Check (Tidy (20.5315) == L"20.5315", "tidy leaves a clean value alone");
	Check (Tidy (static_cast<double> (0.001f)) == L"0.001", "tidy drops float storage noise");
	Check (Tidy (static_cast<double> (0.005f)) == L"0.005", "tidy drops float storage noise 2");
	Check (Tidy (0.07000000000000101) == L"0.07", "tidy drops noise past 14 digits");
	Check (Tidy (-0.3490658503988659) == L"-0.34906585", "tidy keeps a real angle");
	for (double v : { 0.2500000000000021, 1.2345678901234567, 98765.43210987654,
					  0.000123456789012345, -3.14159265358979 })
		{
		double back = 0;
		Check (ParseDouble (Tidy (v), back) && SameDouble (back, v),
			   "tidy stays inside SameDouble");
		}

	for (double v : { 0.1, 0.2 + 0.1, 1.0 / 3.0, 12345.6789, 1e-7, -0.0625,
					  0.05 * 3 })
		{
		double back = 0;
		Check (ParseDouble (FormatDouble (v), back) && back == v,
			   "every printed double parses back to exactly itself");
		}
	}

	// ---- Parsing is strict - a bad cell must fail, not become a number.
	{
	double v = -1;
	Check (ParseDouble (L" 0.25 ", v) && v == 0.25, "surrounding spaces forgiven");
	Check (!ParseDouble (L"", v), "empty cell is not a number");
	Check (!ParseDouble (L"   ", v), "blank cell is not a number");
	Check (!ParseDouble (L"1.5abc", v), "trailing garbage is rejected");
	Check (!ParseDouble (L"1,5", v), "a decimal comma is rejected");
	Check (!ParseDouble (L"NaN", v), "NaN is rejected");
	Check (!ParseDouble (L"inf", v), "infinity is rejected");
	Check (!ParseDouble (L"#REF!", v), "an Excel error is rejected");
	Check (!ParseDouble (L"3-5", v), "a range is rejected");

	long long n = -1;
	Check (ParseLong (L"3", n) && n == 3, "integer parses");
	Check (ParseLong (L"3.0", n) && n == 3, "Excel's 3.0 is still 3");
	Check (!ParseLong (L"3.5", n), "a fraction is not an integer");
	Check (!ParseLong (L"99999999999", n), "out of range is rejected");
	}

	// ---- Booleans.
	{
	bool b = false;
	Check (ParseBool (L"1", b) && b, "1 is true");
	Check (ParseBool (L"0", b) && !b, "0 is false");
	Check (ParseBool (L" TRUE ", b) && b, "TRUE with spaces is true");
	Check (ParseBool (L"False", b) && !b, "False is false");
	Check (ParseBool (L"yes", b) && b, "yes is true");
	Check (!ParseBool (L"maybe", b), "junk is neither");
	Check (!ParseBool (L"", b), "empty is neither");
	}

	// ---- Comparison tolerates a printing round trip and nothing more.
	{
	Check (SameDouble (0.1, 0.1), "equal is same");
	Check (SameDouble (0.1 + 0.2, 0.3), "float dust is same");
	Check (!SameDouble (0.1, 0.11), "a real change is different");
	Check (!SameDouble (0.0, 0.001), "a small real change from zero is different");
	Check (!SameDouble (0.0001, 0.0002), "a tenth-thou change is different");
	Check (!SameDouble (10.0, 10.0001), "a tenth-thou change on a large value is different");
	Check (SameDouble (0.001, static_cast<double> (0.001f)), "float storage noise is same");
	}

	if (gFailed != 0)
		{
		std::printf ("csv_test: %d check(s) FAILED\n", gFailed);
		return 1;
		}
	std::printf ("csv_test: all checks passed\n");
	return 0;
	}
