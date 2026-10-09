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
		bool macros = false;			//!< write an .xlsm with the ribbon tab and its commands
		std::wstring skipKinds;			//!< "|ROUGH|FACE|" - kinds left unticked
		};

	Dump LoadDump ();
	void SaveDump (const Dump &d);

	/// The last sheet dumped, and for which part - where a load starts looking.
	void SetLastDump (const std::wstring &file, const std::wstring &part);
	std::wstring LastDumpFor (const std::wstring &part);

	/// Troubleshooting switches (diag.txt beside the DLL holding a number, else the
	/// DWORD "Diag"; normally neither = 0), to find which
	/// Mastercam call a part objects to: 1 = no Mastercam cycle time,
	/// 2 = no NCI walk (no path stats or estimate).
	unsigned long Diag ();
	}
