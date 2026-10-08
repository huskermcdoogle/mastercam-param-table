//
// MastercamSdk.h - single point of contact with the Mastercam 2026 C-Hook SDK.
//
// stdafx.h already brings in m_core.h and m_mastercam.h (operation / ent structs,
// the operation manager). The two extras are the file manager, which is how the
// open part's name is reached, and the description helper.
//
#pragma once

// The SDK is not warning-clean at /W3 and is not our code to fix.
#pragma warning(push)
#pragma warning(disable : 4267 4244 4996 5054 4458 4275 4251 4471 4838 26495)

#include "m_core.h"
#include "m_mastercam.h"
// m_chookapi.h's IO types use the GRADIENT_* macros from Graphics\GrVars_CH.h and do
// not include it themselves - in 2026 nothing earlier brings it in - so m_graphics.h
// has to come first.
#include "m_graphics.h"
#include "m_chookapi.h"

#pragma warning(pop)
