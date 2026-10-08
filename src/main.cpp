// Mastercam 2026 C++ Add-In: Parameter Table Tool
//
// Two entry points and no UI. Each is its own FUNCTION block in ParamTable.ft,
// so each can sit on a ribbon or the Quick Access Toolbar:
//
//   LatheParamsDumpEntry  - lathe rough / finish / dynamic parameters to CSV
//   LatheParamsLoadEntry  - edited CSVs back into the operations

#include "stdafx.h"
#include "resource.h"
#include "ParamTable.h"
#include "Dump.h"
#include "Load.h"
#include "Util.h"

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
