// Implements the DllMain for the C++ Add-In (MFC extension) DLL

#include "stdafx.h"
#include <afxdllx.h>

/// Handle to the Resources of the C++ Add-In.
HINSTANCE resHandle = nullptr;

/// Retrieves the Handle to the Resources for this DLL.
HINSTANCE GetChookResourceHandle ()
	{
	return resHandle;
	}

static AFX_EXTENSION_MODULE ParamTableDLL = { NULL, NULL };

extern "C" int APIENTRY DllMain (HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
	{
	UNREFERENCED_PARAMETER (lpReserved);

	if (dwReason == DLL_PROCESS_ATTACH)
		{
		TRACE ("ParamTable.DLL Initializing!\n");

		if (!AfxInitExtensionModule (ParamTableDLL, hInstance))
			return 0;

		// Saves the DLL instance handle for later use with ChangeResCl.
		resHandle = hInstance;
		if (resHandle == nullptr)
			return 0;

		ParamTableDLL.hResource = resHandle;
		}
	else if (dwReason == DLL_PROCESS_DETACH)
		{
		TRACE0 ("ParamTable.DLL Terminating!\n");

		// Required because C++ Add-Ins are explicitly loaded by Mastercam.
		if (resHandle)
			AfxFreeLibrary (resHandle);

		AfxTermExtensionModule (ParamTableDLL);
		}

	return 1;
	}
