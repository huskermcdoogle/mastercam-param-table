#pragma once

/// Retrieves the handle to the resources for this C++ Add-In DLL.
///
/// You MUST call -> ChangeResCl res (GetChookResourceHandle ());
/// PRIOR to accessing any "resources" in your C++ Add-In!
HINSTANCE GetChookResourceHandle ();
