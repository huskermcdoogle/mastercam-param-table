//
// ToolPictures.h - a tool's picture as PNG, for the sheet's Tools page.
//
// Mastercam draws it (the same call its setup sheets use) as a BMP; it is
// converted to PNG here, since that is what a workbook embeds well.
// Lathe tools only for now - a mill tool needs a different call.
//
#pragma once

#include <string>

namespace ToolPictures
	{
	/// The lathe tool in this tool-list slot, drawn, as PNG bytes and its size in
	/// pixels. False, with a reason, when there is no such tool or no picture.
	bool LatheTool (long slot, std::string &png, int &width, int &height, std::wstring &why);

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
