//
// DumpDialog.h - the dump's window: which operations, where the sheet goes,
// what it is called, and a couple of options - all remembered for next time.
//
#pragma once

#include "Settings.h"

#include <string>
#include <vector>

namespace DumpDialog
	{
	struct Kind
		{
		std::wstring name;		//!< ROUGH, PRIME, DYNAMIC MILL ...
		int all = 0;			//!< operations of this kind in the part
		int selected = 0;		//!< ... of which selected
		bool on = true;			//!< ticked: dumped
		};

	/// Show it. `kinds` and `selectedOnly` and `settings` come in as the
	/// defaults and go out as chosen. `skipped` is the "not read yet" note.
	/// False = cancelled.
	bool Show (const std::wstring &partFile, std::vector<Kind> &kinds, const std::wstring &skipped,
			   bool &selectedOnly, Settings::Dump &settings);
	}
