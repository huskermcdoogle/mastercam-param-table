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
// The XML is Mastercam's ribbon schema - <Ribbon><Tabs><Tab><Groups><Group>
// <Elements><Button> - each button and the group naming a function table entry as
// "APPLICATION.FUNCTION" exactly as ParamTable.ft declares it (else it resolves to
// nothing, silently). The group is a function too: a HIDDEN one. Learnt on the
// posting app's KinBench (2023); the call says "true" for XML it cannot use, so
// only the tab on screen proves it. What happened goes to ParamTable-ribbon.log.
//
// SDK-free: the Mastercam calls are in main.cpp.
//
#pragma once

#include <string>
#include <vector>

namespace Ribbon
	{
	/// One button: the function table's function, and the command id Mastercam gave
	/// it (to know a tab of the person's own that holds it).
	struct Button
		{
		std::wstring function;		//!< "LatheParamsDumpEntry"
		unsigned id = 0;			//!< from the function table manager, this session
		std::wstring keyTip;		//!< the key after Alt and the tab's own key
		};

	/// The function table APPLICATION in ParamTable.ft - what a button names.
	extern const wchar_t *const kApplication;
	/// The hidden function that names the group.
	extern const wchar_t *const kGroup;

	/// The tab - "Parameter Table", one group of large buttons - as ribbon XML.
	std::wstring TabXml (const std::vector<Button> &buttons);

	/// Whether a workspace file already has any of these commands on a tab of its
	/// own (its <CUSTOM> ribbon sections, buttons by <ID>). The Quick Access Toolbar
	/// does not count - a tab is still worth having.
	bool OnRibbonAlready (const std::string &workspace, const std::vector<unsigned> &ids);
	}
