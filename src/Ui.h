//
// Ui.h - what the tool's own windows (the dump window, the load preview) are
// built on: MFC, WITHOUT the Mastercam SDK.
//
// The windows need nothing of Mastercam but the window to sit over, so they
// include this instead of stdafx.h - and can be shown outside Mastercam too:
// tests\dialog_shots.cpp draws them with sample data for the user manual's
// pictures. That one thing, Mastercam's main window, is Ui::Host - defined with
// the SDK in Util.cpp, and by the picture program (which has no Mastercam).
//
#pragma once

// As stdafx.h has it, so every file of the add-in sees Windows and MFC alike:
// the SDK pins Windows 7 (its BuildTools\targetver_CH.h), and min / max are std's.
#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN
#endif
#include <WinSDKVer.h>
#ifndef _WIN32_WINNT
#define _WIN32_WINNT _WIN32_WINNT_WIN7
#endif
#include <SDKDDKVer.h>
#ifndef _ATL_CSTRING_EXPLICIT_CONSTRUCTORS
#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS
#endif

#include <afxwin.h>				// MFC core
#include <afxcmn.h>				// list, tree and image list controls
#include <afxdlgs.h>			// the folder picker

namespace Ui
	{
	/// The window the tool's windows sit over - Mastercam's main window, so a
	/// window of ours stays in front of it and Mastercam waits while it is open.
	CWnd *Host ();
	}
