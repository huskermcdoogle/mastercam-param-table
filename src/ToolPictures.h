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
	}
