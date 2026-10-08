//
// ColumnHelp.h - one or two plain sentences per sheet column, shown by Excel
// as the cell's tooltip. No Mastercam SDK.
//
#pragma once

#include <string>

namespace ColumnHelp
	{
	/// What the column means, or "" when there is nothing useful to add to
	/// its name. Short: Excel shows at most 255 characters.
	std::wstring For (const std::wstring &name);
	}
