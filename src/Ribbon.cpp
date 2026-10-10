#include "Ribbon.h"

namespace Ribbon
	{
	const wchar_t *const kApplication = L"_ParamTableAddIn_";

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

		/// One CATEGORY as Mastercam's own ribbon writes them (en\RibbonUIRes.dll):
		/// a panel of buttons, each naming its function table function. No image
		/// index - the buttons take the icons ParamTable.ft gives each function.
		std::wstring Category (const std::vector<Button> &buttons)
			{
			std::wstring x = L"<CATEGORY>"
							 L"<ACCTESTING_NAME>ParameterTableTab</ACCTESTING_NAME>"
							 L"<ELEMENT_NAME>Category</ELEMENT_NAME>"
							 L"<NAME>Parameter Table</NAME>"
							 L"<PANELS><PANEL>"
							 L"<ACCTESTING_NAME>ParameterTableGroup</ACCTESTING_NAME>"
							 L"<ELEMENT_NAME>Panel</ELEMENT_NAME>"
							 L"<NAME>Parameter Table</NAME>"
							 L"<ELEMENTS>";
			for (const Button &b : buttons)
				{
				const std::wstring id = std::to_wstring (b.id);
				x += L"<ELEMENT>"
					 L"<ELEMENT_NAME>Button</ELEMENT_NAME>"
					 L"<ID><NAME>" + id + L"</NAME><VALUE>" + id + L"</VALUE>"
					 L"<FT_APPLICATION>" + std::wstring (kApplication) + L"</FT_APPLICATION>"
					 L"<FT_FUNCTION>" + Esc (b.function) + L"</FT_FUNCTION></ID>"
					 L"<TEXT>" + Esc (b.text) + L"</TEXT>"
					 L"<TOOLTIP>" + Esc (b.text) + L"</TOOLTIP>"
					 L"<DESCRIPTION>" + Esc (b.tip) + L"</DESCRIPTION>"
					 L"</ELEMENT>";
				}
			return x + L"</ELEMENTS></PANEL></PANELS></CATEGORY>";
			}
		}

	std::wstring TabXml (const std::vector<Button> &buttons, int shape)
		{
		const std::wstring cat = Category (buttons);
		if (shape == 2)
			return cat;
		const std::wstring cats = L"<CATEGORIES>" + cat + L"</CATEGORIES>";
		if (shape == 1)
			return cats;
		return L"<?xml version=\"1.0\"?><BCGP_RIBBON><HEADER><VERSION>1</VERSION></HEADER>"
			   L"<RIBBON_BAR><ELEMENT_NAME>RibbonBar</ELEMENT_NAME>" + cats + L"</RIBBON_BAR></BCGP_RIBBON>";
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
