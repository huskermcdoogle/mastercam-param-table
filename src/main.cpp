// Mastercam 2026 C++ Add-In: Parameter Table Tool
//
// Five entry points. Each is its own FUNCTION block in ParamTable.ft, so each
// can sit on a ribbon or the Quick Access Toolbar:
//
//   LatheParamsDumpEntry  - the operations' parameters to a workbook
//   LatheParamsLoadEntry  - the edited workbook back into the operations
//   LatheParamsUndoEntry  - the last load's old values back, from ParamTable.log
//   LatheParamsRegenEntry - regenerate what the last load changed, asked first
//   ParamTableHelpEntry   - the user manual (help\index.html beside the DLL)

#include "stdafx.h"
#include "resource.h"
#include "ParamTable.h"
#include "Dump.h"
#include "Load.h"
#include "Regen.h"
#include "Settings.h"
#include "Util.h"

#include "Ribbon.h"

#include <shellapi.h>

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <set>
#include <sstream>

namespace
	{
	/// ParamTable-ribbon.log beside the DLL: what the startup did about the ribbon -
	/// Mastercam documents neither when an add-in hears from it nor the ribbon XML it
	/// takes, so this is how a real start says. Kept short: started again past 64 KB.
	void RibbonLog (const std::wstring &line)
		{
		const std::wstring folder = Settings::AddinFolder ();
		if (folder.empty ())
			return;
		const std::filesystem::path file (folder + L"ParamTable-ribbon.log");
		std::error_code ec;
		const bool big = std::filesystem::exists (file, ec) && std::filesystem::file_size (file, ec) > 64 * 1024;
		std::ofstream f (file, big ? std::ios::trunc : std::ios::app);
		SYSTEMTIME t;
		GetLocalTime (&t);
		wchar_t stamp[32];
		swprintf_s (stamp, L"%02u:%02u:%02u  ", t.wHour, t.wMinute, t.wSecond);
		const std::wstring w = stamp + line + L"\n";
		const int n = WideCharToMultiByte (CP_UTF8, 0, w.c_str (), static_cast<int> (w.size ()), nullptr, 0, nullptr, nullptr);
		std::string u (static_cast<size_t> (n), '\0');
		WideCharToMultiByte (CP_UTF8, 0, w.c_str (), static_cast<int> (w.size ()), u.data (), n, nullptr, nullptr);
		f << u;
		}

	/// Every workspace file in Mastercam's CONFIG folder, read whole - where the tabs a
	/// person made with Customize are kept.
	std::string Workspaces ()
		{
		wchar_t dir[MAX_PATH] = {};
		DWORD size = sizeof (dir);
		if (RegGetValueW (HKEY_CURRENT_USER, L"SOFTWARE\\CNC Software\\Mastercam 2026", L"UserDir", RRF_RT_REG_SZ,
						  nullptr, dir, &size) != ERROR_SUCCESS)
			return std::string ();
		std::string all;
		std::error_code ec;
		for (const auto &e : std::filesystem::directory_iterator (std::filesystem::path (dir) / L"CONFIG", ec))
			if (_wcsicmp (e.path ().extension ().c_str (), L".Workspace") == 0)
				{
				std::ifstream f (e.path (), std::ios::binary);
				std::ostringstream s;
				s << f.rdbuf ();
				all += s.str ();
				}
		return all;
		}

	/// The add-in's tab on the ribbon, once a session - unless its commands are on a
	/// tab already (one made with Customize keeps its place, no second copy).
	void AddRibbonTab (const wchar_t *when)
		{
		static bool done = false;
		if (done)
			return;
		done = true;
		std::vector<Ribbon::Button> buttons = {
			{ L"LatheParamsDumpEntry", 0, L"D" }, { L"LatheParamsLoadEntry", 0, L"L" },
			{ L"LatheParamsUndoEntry", 0, L"U" }, { L"LatheParamsRegenEntry", 0, L"R" },
			{ L"ParamTableHelpEntry", 0, L"H" } };
		std::vector<unsigned> ids;
		std::wstring said = std::wstring (L"ribbon (") + when + L"): command ids";
		const Cnc::IFunctionTableManagerPtr ft = Cnc::GetFunctionTableManager ();
		for (Ribbon::Button &b : buttons)
			{
			b.id = ft ? ft->GetCommandIdByName (Ribbon::kApplication, b.function.c_str ()) : 0;
			ids.push_back (b.id);
			said += L" " + b.function + L"=" + std::to_wstring (b.id);
			}
		RibbonLog (said);
		if (std::find (ids.begin (), ids.end (), 0u) != ids.end ())
			{
			RibbonLog (L"ribbon: a command has no id - the function table is not loaded yet; no tab");
			return;
			}
		if (Ribbon::OnRibbonAlready (Workspaces (), ids))
			{
			RibbonLog (L"ribbon: the commands are on a tab of your own already (Customize) - no second tab");
			return;
			}
		// A marker while the tab is offered: a start that finds it left over is one after
		// an offer that never came back (Mastercam went down with it) - so it is not
		// offered again, and Mastercam starts. Delete the marker to try again.
		const std::filesystem::path marker (Settings::AddinFolder () + L"ParamTable-ribbon.trying");
		std::error_code ec;
		if (std::filesystem::exists (marker, ec))
			{
			RibbonLog (L"ribbon: the last offer never finished - not offered again (delete ParamTable-ribbon.trying to retry)");
			return;
			}
		std::ofstream (marker) << "offering the ribbon tab\n";
		RibbonLog (L"ribbon: offering the tab");
		const bool ok = InsertThirdPartyRibbonTabs (Ribbon::TabXml (buttons).c_str ());
		RibbonLog (ok ? L"ribbon: taken (only the tab on screen proves it worked)"
					  : L"ribbon: refused - add the commands with Customize instead");
		std::filesystem::remove (marker, ec);
		}
	}

extern "C" __declspec(dllexport) int m_open (int not_used)
	{
	// Loaded at a start - and again for each command (they unload it when done).
	RibbonLog (L"m_open: Mastercam loaded the add-in");
	return MC_NOERROR;
	}

extern "C" __declspec(dllexport) int m_close (int not_used)
	{
	return MC_NOERROR;
	}

/// Mastercam's events. Ready (MCEVENT_READY) is when the ribbon can take a tab.
/// Each kind of event is logged once, so a start shows what an add-in hears.
extern "C" __declspec(dllexport) int m_notify (int notify_code)
	{
	static std::set<int> heard;
	if (heard.size () < 40 && heard.insert (notify_code).second)
		RibbonLog (L"m_notify: event " + std::to_wstring (notify_code));
	if (notify_code == MCEVENT_READY)
		{
		try
			{
			AddRibbonTab (L"ready");
			}
		catch (...)
			{
			RibbonLog (L"ribbon: failed");
			}
		}
	return MC_NOERROR;
	}

/// The ribbon tab's group. The ribbon names a group by a function, like a button,
/// so it is one (HIDDEN in ParamTable.ft) - and does nothing.
extern "C" __declspec(dllexport) int ParamTableGroup (int param)
	{
	return MC_NOERROR;
	}

extern "C" __declspec(dllexport) int LatheParamsDumpEntry (int param)
	{
	// Must call this prior to accessing any Resources in the C++ Add-In DLL !
	ChangeResCl res (GetChookResourceHandle ());

	try
		{
		Dump::Run ();
		}
	catch (...)
		{
		Util::Say (L"The dump failed unexpectedly. Nothing was changed in the part.",
				   MB_ICONERROR);
		}
	return MC_NOERROR | MC_UNLOADAPP;
	}

extern "C" __declspec(dllexport) int LatheParamsLoadEntry (int param)
	{
	ChangeResCl res (GetChookResourceHandle ());

	try
		{
		Load::Run ();
		}
	catch (...)
		{
		// A throw here can land partway through a write, so it is said plainly
		// rather than swallowed - the log is what says how far it got.
		Util::Say (L"The load failed unexpectedly. Check ParamTable.log beside "
				   L"the part to see which operations were written before it "
				   L"stopped.", MB_ICONERROR);
		}
	return MC_NOERROR | MC_UNLOADAPP;
	}

/// Its own command rather than a button in the load: a load is usually judged
/// AFTER it - regenerated, backplotted, posted - when the load's window is long
/// closed. Beside "load" on a ribbon, it is there when that moment comes.
extern "C" __declspec(dllexport) int LatheParamsUndoEntry (int param)
	{
	ChangeResCl res (GetChookResourceHandle ());

	try
		{
		Load::UndoLast ();
		}
	catch (...)
		{
		Util::Say (L"The undo failed unexpectedly. Check ParamTable.log beside "
				   L"the part to see which operations were restored before it "
				   L"stopped.", MB_ICONERROR);
		}
	return MC_NOERROR | MC_UNLOADAPP;
	}

/// Regenerate the operations the last load changed - on purpose, never on the
/// side: on a big part it takes twenty minutes. It asks first (No is the
/// default), then puts each one's Mastercam cycle time beside the estimate the
/// load's sheet showed, in the part's history.
extern "C" __declspec(dllexport) int LatheParamsRegenEntry (int param)
	{
	ChangeResCl res (GetChookResourceHandle ());

	try
		{
		Regen::LastLoad ();
		}
	catch (...)
		{
		Util::Say (L"The regeneration failed unexpectedly. Check ParamTable.log beside the part to see which "
				   L"operations were regenerated before it stopped.", MB_ICONERROR);
		}
	return MC_NOERROR | MC_UNLOADAPP;
	}

/// The user manual: a folder of plain web pages that ships beside the DLL
/// (Add-Ins\ParamTable\help). Opened in the default browser - it needs no
/// internet, and nothing in Mastercam changes.
extern "C" __declspec(dllexport) int ParamTableHelpEntry (int param)
	{
	ChangeResCl res (GetChookResourceHandle ());

	try
		{
		const std::wstring folder = Settings::AddinFolder ();
		const std::wstring page = folder + L"help\\index.html";
		if (folder.empty () || GetFileAttributesW (page.c_str ()) == INVALID_FILE_ATTRIBUTES)
			{
			Util::Say (L"The user manual is not installed. It comes with the add-in, in the help "
					   L"folder beside it:\r\n\r\n  " + (folder.empty () ? std::wstring (L"Add-Ins\\ParamTable\\help") : folder + L"help")
					   + L"\r\n\r\nCopy the whole ParamTable folder from the release zip again.");
			return MC_NOERROR | MC_UNLOADAPP;
			}
		const HINSTANCE h = ShellExecuteW (nullptr, L"open", page.c_str (), nullptr, nullptr, SW_SHOWNORMAL);
		if (reinterpret_cast<INT_PTR> (h) <= 32)
			Util::Say (L"Windows could not open the user manual in a browser. Open this file by hand:"
					   L"\r\n\r\n  " + page, MB_ICONWARNING);
		}
	catch (...)
		{
		Util::Say (L"Could not open the user manual.", MB_ICONERROR);
		}
	return MC_NOERROR | MC_UNLOADAPP;
	}

extern "C" __declspec(dllexport) int m_version (int version)
	{
	int ret = C_H_VERSION;

	// Any release of the same major version.
	if ((version / 100) == (C_H_VERSION / 100))
		ret = version;

	return ret;
	}

/// Run from Settings > Run User Application, with no function table involved:
/// asks which of the two to do.
extern "C" __declspec(dllexport) int m_main (int not_used)
	{
	ChangeResCl res (GetChookResourceHandle ());

	const int a = Util::Say (L"Dump the lathe parameters to Excel?\r\n\r\n"
							 L"Yes = dump     No = load an edited sheet     Cancel = nothing",
							 MB_YESNOCANCEL | MB_ICONQUESTION);
	if (a == IDYES)
		return LatheParamsDumpEntry (0);
	if (a == IDNO)
		return LatheParamsLoadEntry (0);
	return MC_NOERROR | MC_UNLOADAPP;
	}
