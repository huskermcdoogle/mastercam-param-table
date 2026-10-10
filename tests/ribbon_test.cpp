// The add-in's own ribbon tab: the XML it hands Mastercam (in each shape it offers),
// and whether the commands are on a tab of the person's own already.
#include "../src/Ribbon.h"

#include <cstdio>

static int failed = 0;
static void Check (bool ok, const char *what)
	{
	std::printf ("  %s  %s\n", ok ? "ok  " : "FAIL", what);
	if (!ok)
		++failed;
	}

static size_t Count (const std::wstring &s, const std::wstring &what)
	{
	size_t n = 0;
	for (size_t at = s.find (what); at != std::wstring::npos; at = s.find (what, at + 1))
		++n;
	return n;
	}

int main ()
	{
	const std::vector<Ribbon::Button> buttons = {
		{ L"LatheParamsDumpEntry", 42014, L"Dump to Excel", L"Write the parameters." },
		{ L"LatheParamsLoadEntry", 42015, L"Load from Excel", L"Load <edited> & checked." } };

	const std::wstring whole = Ribbon::TabXml (buttons, 0);
	Check (whole.rfind (L"<?xml", 0) == 0 && whole.find (L"<BCGP_RIBBON>") != std::wstring::npos
			   && whole.find (L"</BCGP_RIBBON>") != std::wstring::npos, "shape 0: a whole BCGP_RIBBON document");
	Check (Count (whole, L"<CATEGORY>") == 1 && Count (whole, L"<ELEMENT>") == 2, "one tab, two buttons");
	Check (whole.find (L"<ID><NAME>42014</NAME><VALUE>42014</VALUE><FT_APPLICATION>_ParamTableAddIn_</FT_APPLICATION>"
					   L"<FT_FUNCTION>LatheParamsDumpEntry</FT_FUNCTION></ID>") != std::wstring::npos,
		   "a button names its command id and its function table function, as Mastercam's own do");
	Check (whole.find (L"Load &lt;edited&gt; &amp; checked.") != std::wstring::npos, "text is escaped for XML");
	for (const wchar_t *tag : { L"CATEGORY", L"PANELS", L"PANEL", L"ELEMENTS", L"ELEMENT", L"ID", L"CATEGORIES", L"RIBBON_BAR" })
		{
		const std::wstring open = std::wstring (L"<") + tag + L">", close = std::wstring (L"</") + tag + L">";
		if (Count (whole, open) != Count (whole, close))
			Check (false, "every tag closed");
		}
	const std::wstring cats = Ribbon::TabXml (buttons, 1), cat = Ribbon::TabXml (buttons, 2);
	Check (cats.rfind (L"<CATEGORIES><CATEGORY>", 0) == 0 && cat.rfind (L"<CATEGORY>", 0) == 0
			   && whole.find (cats) != std::wstring::npos, "shapes 1 and 2: the CATEGORIES, the CATEGORY - the same tab");

	// A workspace as Mastercam writes one: a tab made with Customize, and the Quick
	// Access Toolbar.
	const std::string custom = "<CUSTOM><CATEGORY><KEY>1024</KEY><TEXT>Parameter Table Tool</TEXT><PANELS><PANEL>"
							   "<ELEMENTS><ELEMENT><ID>42014</ID><TEXT>Lathe params - dump to Excel</TEXT></ELEMENT>"
							   "</ELEMENTS></PANEL></PANELS></CATEGORY></CUSTOM>";
	const std::string qat = "<QAT_ELEMENTS><ITEM><ID>42014</ID></ITEM></QAT_ELEMENTS>";
	Check (Ribbon::OnRibbonAlready ("<GUI>" + custom + "</GUI>", { 42014, 42015 }), "on a tab of their own: found");
	Check (!Ribbon::OnRibbonAlready ("<GUI>" + qat + "</GUI>", { 42014, 42015 }), "only on the Quick Access Toolbar: still add the tab");
	Check (!Ribbon::OnRibbonAlready ("<GUI>" + custom + "</GUI>", { 52014 }), "other commands on a custom tab: not ours");
	Check (!Ribbon::OnRibbonAlready ("<CUSTOM><ID>420140</ID></CUSTOM>", { 42014 }), "an id is matched whole, not a part of a longer one");
	Check (!Ribbon::OnRibbonAlready ("", { 42014 }), "no workspace: add the tab");

	if (failed)
		std::printf ("ribbon_test: %d FAILED\n", failed);
	else
		std::printf ("ribbon_test: all checks passed\n");
	return failed ? 1 : 0;
	}
