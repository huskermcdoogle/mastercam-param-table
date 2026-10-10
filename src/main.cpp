// Mastercam 2026 C++ Add-In: Parameter Table Tool
//
// Four entry points. Each is its own FUNCTION block in ParamTable.ft, so each
// can sit on a ribbon or the Quick Access Toolbar:
//
//   LatheParamsDumpEntry  - the operations' parameters to a workbook
//   LatheParamsLoadEntry  - the edited workbook back into the operations
//   LatheParamsUndoEntry  - the last load's old values back, from ParamTable.log
//   ParamTableHelpEntry   - the user manual (help\index.html beside the DLL)

#include "stdafx.h"
#include "resource.h"
#include "ParamTable.h"
#include "Dump.h"
#include "Load.h"
#include "Settings.h"
#include "Util.h"

#include <shellapi.h>

extern "C" __declspec(dllexport) int m_open (int not_used)
	{
	return MC_NOERROR;
	}

extern "C" __declspec(dllexport) int m_close (int not_used)
	{
	return MC_NOERROR;
	}

extern "C" __declspec(dllexport) int m_notify (int notify_code)
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
