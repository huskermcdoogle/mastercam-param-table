//
// Ribbon.h - the add-in's own tab on Mastercam's ribbon, so a new install needs no
// trip to Customize: when Mastercam says it is ready (m_notify, MCEVENT_READY), the
// add-in hands it a tab - "Parameter Table" - with its four commands
// (InsertThirdPartyRibbonTabs, Mastercam's own ribbon XML).
//
// NOT when the commands are on a ribbon already: someone who put them on a tab of
// their own (Customize) keeps that, and gets no second copy. The workspace file
// holds those custom tabs, each button by its command id.
//
// Mastercam documents the call but not the XML it wants, so the tab is offered in
// the shapes its own ribbon XML suggests, most likely first, until one is taken;
// what happened goes to ParamTable-ribbon.log beside the DLL.
//
// SDK-free: the Mastercam calls are in main.cpp.
//
#pragma once

#include <string>
#include <vector>

namespace Ribbon
	{
	/// One button: the function table's function, and the command id Mastercam gave it.
	struct Button
		{
		std::wstring function;		//!< "LatheParamsDumpEntry"
		unsigned id = 0;			//!< from the function table manager, this session
		std::wstring text;			//!< under the button
		std::wstring tip;			//!< its tooltip's description
		};

	/// The function table APPLICATION in ParamTable.ft - what a button names.
	extern const wchar_t *const kApplication;

	/// How many shapes TabXml has to offer.
	const int kShapes = 3;

	/// The tab as Mastercam's ribbon XML, in one of its shapes: 0 a whole
	/// BCGP_RIBBON document, 1 its CATEGORIES, 2 the CATEGORY alone.
	std::wstring TabXml (const std::vector<Button> &buttons, int shape);

	/// Whether a workspace file already has any of these commands on a tab of its
	/// own (its <CUSTOM> ribbon sections, buttons by <ID>). The Quick Access Toolbar
	/// does not count - a tab is still worth having.
	bool OnRibbonAlready (const std::string &workspace, const std::vector<unsigned> &ids);
	}
