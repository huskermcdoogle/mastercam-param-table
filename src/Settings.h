//
// Settings.h - what the tool remembers between runs, per Windows user:
// HKEY_CURRENT_USER\Software\ParamTableTool. Nothing here is needed - every
// value has a default, and a missing or unreadable key just means defaults.
//
#pragma once

#include <string>
#include <utility>
#include <vector>

namespace Settings
	{
	struct Dump
		{
		std::wstring folder;			//!< "" = beside the part
		std::wstring pattern;			//!< file name pattern - see FileRules.h
		bool openExcel = true;
		bool pictures = true;
		bool macros = false;			//!< write an .xlsm with the ribbon tab and its commands
		bool stockSim = true;			//!< simulate the stock for removed volume (else Mastercam's boundaries)
		std::wstring skipKinds;			//!< "|ROUGH|FACE|" - kinds left unticked
		};

	Dump LoadDump ();
	void SaveDump (const Dump &d);

	/// The last sheet dumped, and for which part - where a load starts looking.
	void SetLastDump (const std::wstring &file, const std::wstring &part);
	std::wstring LastDumpFor (const std::wstring &part);

	/// SAVED SELECTIONS: named sets of ticked operations (by op number), per part
	/// file name - Selections\<part file name>, one value per name, "2,5,12".
	/// In name order.
	std::vector<std::pair<std::wstring, std::vector<long>>> Selections (const std::wstring &partFile);
	void SaveSelection (const std::wstring &partFile, const std::wstring &name, const std::vector<long> &ids);
	void DeleteSelection (const std::wstring &partFile, const std::wstring &name);

	/// Troubleshooting switches (diag.txt beside the DLL holding a number, else the
	/// DWORD "Diag"; normally neither = 0), to find which
	/// Mastercam call a part objects to: 1 = no Mastercam cycle time,
	/// 2 = no NCI walk (no path stats or estimate), 4 = stock sim probe (removed
	/// volume per op from a 2D stock simulation, logged only - StockSim.h).
	unsigned long Diag ();

	/// The add-in's own folder (where ParamTable.dll is), with a trailing "\\".
	std::wstring AddinFolder ();
	}
