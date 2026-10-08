// File naming rules for a dump.
#include "../src/FileRules.h"

#include <cstdio>
#include <fstream>

int main (int argc, char **argv)
	{
	int failed = 0;
	auto check = [&failed] (bool ok, const char *what)
		{
		if (!ok)
			{
			std::printf ("FAIL  %s\n", what);
			++failed;
			}
		};
	std::tm tmv = {};
	tmv.tm_year = 2026 - 1900; tmv.tm_mon = 9; tmv.tm_mday = 8; tmv.tm_hour = 14; tmv.tm_min = 5; tmv.tm_sec = 9; tmv.tm_isdst = -1;
	const std::time_t when = std::mktime (&tmv);

	using FileRules::Name;
	check (Name (FileRules::kDefaultPattern, L"Sample Part", when, false, 15)
		   == L"Sample Part_lathe_params_20261008-140509.xlsx", "default pattern");
	check (Name (L"{part} {scope} {ops} ops", L"P1", when, true, 3) == L"P1 selected 3 ops.xlsx", "scope and ops");
	check (Name (L"", L"P1", when, false, 1) == L"P1_lathe_params_20261008-140509.xlsx", "empty pattern = default");
	check (Name (L"a/b:c*?\"<>|", L"P", when, false, 1) == L"a_b_c______.xlsx", "forbidden characters become _");
	check (Name (L"{nope} x", L"P", when, false, 1) == L"{nope} x.xlsx", "unknown token kept as typed");
	check (Name (L"name.xlsx", L"P", when, false, 1) == L"name.xlsx", ".xlsx not doubled");
	check (Name (L"trailing. ", L"P", when, false, 1) == L"trailing.xlsx", "trailing dots and spaces dropped");

	// Never overwrite.
	const std::filesystem::path dir = std::filesystem::path (argc > 1 ? argv[1] : ".") / "filerules";
	std::filesystem::create_directories (dir);
	std::filesystem::remove (dir / "x.xlsx");
	std::filesystem::remove (dir / "x (2).xlsx");
	check (FileRules::Unique (dir, L"x.xlsx") == dir / "x.xlsx", "free name used as is");
	{ std::ofstream (dir / "x.xlsx") << "1"; }
	check (FileRules::Unique (dir, L"x.xlsx") == dir / "x (2).xlsx", "taken name gets (2)");
	{ std::ofstream (dir / "x (2).xlsx") << "1"; }
	check (FileRules::Unique (dir, L"x.xlsx") == dir / "x (3).xlsx", "then (3)");

	if (failed)
		{
		std::printf ("filerules_test: %d check(s) FAILED\n", failed);
		return 1;
		}
	std::printf ("filerules_test: all checks passed\n");
	return 0;
	}
