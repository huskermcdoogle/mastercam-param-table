//
// Util.h - the small things every entry point needs: where the open part is,
// what to tell the operator, and the log that makes a reload undoable.
//
#pragma once

#include <filesystem>
#include <string>

namespace Util
	{
	/// The open part's .mcam path, or empty when nothing is saved yet.
	std::filesystem::path PartFile ();

	/// A message box parented to Mastercam's main window.
	int Say (const std::wstring &text, unsigned flags = 0x40 /*MB_ICONINFORMATION*/);

	/// Append a line to ParamTable.log beside the part, time-stamped. Every
	/// value a reload overwrites goes here first - it is the way back, and the
	/// record "undo last load" reads (Undo.h).
	void Log (const std::filesystem::path &partFile, const std::wstring &line);
	}
