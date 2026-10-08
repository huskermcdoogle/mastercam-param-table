// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently

#pragma once

#define NOMINMAX    // Remove the min and max definitions in a Windows header to avoid conflicts with std::min/std::max.
#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // Exclude rarely-used stuff from Windows headers
#endif

#include "3DVarsDefines_CH.h" // for C_H_VERSION

#if C_H_VERSION >= 2800 // Mastercam 2026 or later
#undef _WIN32_WINNT
#include "BuildTools/targetver_CH.h"
#else
#include "targetver.h"
#endif

#if C_H_VERSION >= 2100 // Mastercam 2019 or later
#undef _MBCS
#ifndef _UNICODE
#define _UNICODE
#endif
//#pragma message("Building for Mastercam with _UNICODE Character Set")
#else // Mastercam prior to 2019
#undef _UNICODE
#ifndef _MBCS
#define _MBCS
#endif
//#pragma message("Building for Mastercam _MBCS Character Set")
#endif

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS  // some CString constructors will be explicit

#include <afxwin.h>             // MFC core and standard components
#include <afxext.h>             // MFC extensions

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxole.h>             // MFC OLE classes
#include <afxodlgs.h>           // MFC OLE dialog classes
#include <afxdisp.h>            // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT

#ifndef _AFX_NO_DB_SUPPORT
#include <afxdb.h>              // MFC ODBC database classes
#endif // _AFX_NO_DB_SUPPORT

#ifndef _AFX_NO_DAO_SUPPORT
#include <afxdao.h>             // MFC DAO database classes
#endif // _AFX_NO_DAO_SUPPORT

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC support for Internet Explorer 4 Common Controls
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT
#include <afxdlgs.h>            // MFC common dialogs (CFileDialog) - the load picks its files with one

// Mastercam specific items:
// These standard includes are required here starting with Mastercam 2018
#include <array>
#include <deque>
#include <functional>
#include <map>
#include <memory>
#include <set>
#include <unordered_map>
#include <unordered_set>
#include <optional>
#include <filesystem>

// This will cause the proper BCGCBPRO#######.lib to be linked in.
#include "BCGCBProInc.h"

// Include these "main" Mastercam header files.
#include "m_core.h"
#include "m_mastercam.h"
#if C_H_VERSION >= 2200 // Mastercam 2020 or later
#include "m_corebase.h"
#include "m_guibase.h"
#endif
#if C_H_VERSION >= 2800 // Mastercam 2020 or later
#include "OpmanGUI_CH.h"
#endif

// Optional 'master' headers that include groups of related headers ->
// "m_CHookAPI.h" // Mastercam 2022 or later
// m_msurf.h
// m_post.h
// m_solids.h
// m_tool.h
// m_wire.h
// m_bnci.h
// m_cad.h
// m_cadutil.h
// m_control.h
// m_curvesurfgeom.h
// m_lathe.h
// m_machinedef.h
// m_math.h
// m_mill.h
