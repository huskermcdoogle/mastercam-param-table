//
// Settings.h - what the tool remembers between runs, per Windows user:
// HKEY_CURRENT_USER\Software\ParamTableTool. Nothing here is needed - every
// value has a default, and a missing or unreadable key just means defaults.
//
#pragma once

#include <string>

namespace Settings
	{
	struct Dump
		{
		std::wstring folder;			//!< "" = beside the part
		std::wstring pattern;			//!< file name pattern - see FileRules.h
		bool openExcel = true;
		bool pictures = true;
		std::wstring skipKinds;			//!< "|ROUGH|FACE|" - kinds left unticked
		};

	Dump LoadDump ();
	void SaveDump (const Dump &d);

	/// The last sheet dumped, and for which part - where a load starts looking.
	void SetLastDump (const std::wstring &file, const std::wstring &part);
	std::wstring LastDumpFor (const std::wstring &part);
	}
