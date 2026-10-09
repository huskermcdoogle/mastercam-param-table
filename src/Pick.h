//
// Pick.h - the dump window's search and saved selections, without the window.
// No Mastercam SDK, no MFC.
//
#pragma once

#include <string>
#include <vector>

namespace Pick
	{
	/// Whether an operation's line ("op 12    ROUGH    T3    Rough OD") matches
	/// what was typed: every word of it appears somewhere in the line, in any
	/// case and any order - "t3 rough" finds the roughs on tool 3. Nothing typed
	/// matches nothing (no search is running).
	bool Matches (const std::wstring &line, const std::wstring &typed);

	/// Op numbers as saved: "2,5,12".
	std::wstring IdsText (const std::vector<long> &ids);

	/// And back. Anything that is not a number is skipped.
	std::vector<long> ParseIds (const std::wstring &text);

	/// What the selections of a part are filed under: its file name, lower case
	/// (the same part from another folder keeps its selections).
	std::wstring PartKey (const std::wstring &partFile);

	/// A selection name as stored: trimmed, at most 60 characters. Empty = not
	/// a usable name.
	std::wstring CleanName (const std::wstring &name);
	}
