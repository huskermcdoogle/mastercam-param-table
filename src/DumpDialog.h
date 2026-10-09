//
// DumpDialog.h - the dump's window: which operations, where the sheet goes,
// what it is called, and a few options - all but the operations remembered.
//
#pragma once

#include "Settings.h"

#include <string>
#include <vector>

namespace DumpDialog
	{
	/// One operation the tool can dump, as the window lists it.
	struct Op
		{
		long idn = 0;
		long group = 0;			//!< its toolpath group
		std::wstring groupName;
		std::wstring kind;		//!< ROUGH, PRIME, DYNAMIC MILL ...
		std::wstring text;		//!< "op 12   ROUGH   T3   Rough OD"
		bool selectedInMgr = false;	//!< selected in the Operation Manager
		bool on = false;		//!< ticked: dumped
		};

	/// Show it. `ops` (in Operation Manager order) and `settings` come in as the
	/// defaults and go out as chosen. `skipped` is the "not read yet" note.
	/// False = cancelled.
	bool Show (const std::wstring &partFile, std::vector<Op> &ops, const std::wstring &skipped,
			   Settings::Dump &settings);
	}
