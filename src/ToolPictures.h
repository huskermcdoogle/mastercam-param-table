//
// ToolPictures.h - a tool's picture as PNG, for the sheet's Tools page.
//
// Mastercam draws it as a bitmap - a lathe tool through the call its setup
// sheets use, a mill tool through the tool manager's bitmap factory - and it
// is converted to PNG here, since that is what a workbook embeds well.
//
#pragma once

#include <string>

namespace ToolPictures
	{
	/// The lathe tool in this tool-list slot, drawn, as PNG bytes and its size in
	/// pixels. False, with a reason, when there is no such tool or no picture.
	bool LatheTool (long slot, std::string &png, int &width, int &height, std::wstring &why);

	/// The same for the MILL tool in this mill tool-list slot. Never throws and
	/// never asks anything: a failure is false with a reason.
	bool MillTool (long slot, std::string &png, int &width, int &height, std::wstring &why);

	/// The insert name of the lathe tool in this slot ("" when none).
	std::wstring LatheInsert (long slot);

	/// The lathe tool's manufacturer code (a 3D tool's order code; "" when blank).
	std::wstring LatheMfgCode (long slot);

	/// The lathe tool's main insert as the tool manager defines it.
	struct InsertInfo
		{
		bool ok = false;			//!< the tool has an insert definition
		wchar_t shape = 0;			//!< ANSI shape code: C, D, V, T, S, R, W ... (0 = none / custom)
		double ic = 0, radius = 0, thickness = 0, width = 0, length = 0;	//!< in the part's units
		std::wstring grade;			//!< the insert grade's name
		bool custom = false;
		};
	InsertInfo LatheInsertInfo (long slot);
	}
