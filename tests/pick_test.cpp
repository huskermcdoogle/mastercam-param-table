// The dump window's search and saved selections, without the window.
#include "../src/Pick.h"

#include <cstdio>

namespace
	{
	int failed = 0;

	void Check (bool ok, const char *what)
		{
		std::printf ("  %s  %s\n", ok ? "ok   " : "FAIL ", what);
		if (!ok)
			++failed;
		}
	}

int main ()
	{
	const std::wstring line = L"op 12    ROUGH    T3    Rough OD - 1st side";
	Check (Pick::Matches (line, L"rough"), "a word, any case");
	Check (Pick::Matches (line, L"  t3   ROUGH "), "several words, any order, extra spaces");
	Check (!Pick::Matches (line, L"t3 finish"), "every word must be there");
	Check (Pick::Matches (line, L"op 12"), "the op number");
	Check (Pick::Matches (line, L"1st"), "the comment");
	Check (!Pick::Matches (line, L""), "nothing typed matches nothing");
	Check (!Pick::Matches (line, L"   "), "only spaces matches nothing");

	Check (Pick::IdsText ({ 2, 5, 12 }) == L"2,5,12", "ids to text");
	Check (Pick::IdsText ({}) == L"", "no ids");
	const std::vector<long> ids = Pick::ParseIds (L"2,5, 12,x,,40");
	Check (ids.size () == 4 && ids[0] == 2 && ids[2] == 12 && ids[3] == 40, "text to ids, junk skipped");
	Check (Pick::ParseIds (L"").empty (), "empty text, no ids");
	Check (Pick::ParseIds (L"99999999999999999999").size () == 3, "a huge number does not overflow");

	Check (Pick::PartKey (L"C:\\Parts\\Oil Spool.MCAM") == L"oil spool.mcam", "the part's key: its file name, lower case");
	Check (Pick::PartKey (L"plain.mcam") == L"plain.mcam", "a bare name");
	Check (Pick::CleanName (L"  finish ops \t") == L"finish ops", "a name trimmed");
	Check (Pick::CleanName (std::wstring (100, L'x')).size () == 60, "a name kept to 60 characters");

	if (failed)
		{
		std::printf ("pick_test: %d FAILED\n", failed);
		return 1;
		}
	std::puts ("pick_test: all checks passed");
	return 0;
	}
