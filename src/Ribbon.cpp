#include "Ribbon.h"

namespace Ribbon
	{
	const wchar_t *const kApplication = L"_ParamTableAddIn_";
	const wchar_t *const kGroup = L"ParamTableGroup";

	namespace
		{
		std::wstring Esc (const std::wstring &s)
			{
			std::wstring o;
			for (wchar_t c : s)
				switch (c)
					{
					case L'&': o += L"&amp;"; break;
					case L'<': o += L"&lt;"; break;
					case L'>': o += L"&gt;"; break;
					case L'"': o += L"&quot;"; break;
					default: o += c;
					}
			return o;
			}

		/// "APPLICATION.FUNCTION", as the ribbon names a function table entry.
		std::wstring Command (const std::wstring &function)
			{
			return Esc (std::wstring (kApplication) + L"." + function);
			}
		}

	std::wstring TabXml (const std::vector<Button> &buttons)
		{
		// Labels, icons and tips come from the function table entries themselves.
		std::wstring x = L"<?xml version=\"1.0\" encoding=\"utf-8\"?>"
						 L"<Ribbon><Tabs><Tab Label=\"Parameter Table\" KeyTip=\"PT\"><Groups>"
						 L"<Group FTCommand=\"" + Command (kGroup) + L"\" KeyTip=\"PG\"><Elements>";
		for (const Button &b : buttons)
			x += L"<Button FTCommand=\"" + Command (b.function) + L"\" KeyTip=\"" + Esc (b.keyTip)
				 + L"\" LargeMode=\"Always\" />";
		return x + L"</Elements></Group></Groups></Tab></Tabs></Ribbon>";
		}

	bool OnRibbonAlready (const std::string &workspace, const std::vector<unsigned> &ids)
		{
		// Only inside <CUSTOM> ... </CUSTOM>: the tabs a person made. The same ids
		// elsewhere (the Quick Access Toolbar, key maps) are not a tab.
		for (size_t at = workspace.find ("<CUSTOM>"); at != std::string::npos; at = workspace.find ("<CUSTOM>", at + 1))
			{
			const size_t end = workspace.find ("</CUSTOM>", at);
			const std::string part = workspace.substr (at, end == std::string::npos ? std::string::npos : end - at);
			for (unsigned id : ids)
				if (id != 0 && part.find ("<ID>" + std::to_string (id) + "</ID>") != std::string::npos)
					return true;
			}
		return false;
		}
	}
