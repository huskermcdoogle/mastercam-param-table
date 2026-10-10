// The add-in's own ribbon tab: the XML it hands Mastercam,
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
		{ L"LatheParamsDumpEntry", 42014, L"D" }, { L"LatheParamsLoadEntry", 42015, L"L" } };

	// Mastercam's ribbon schema, as the posting app's KinBench tab used it.
	const std::wstring x = Ribbon::TabXml (buttons);
	Check (x.rfind (L"<?xml", 0) == 0 && x.find (L"<Ribbon><Tabs><Tab Label=\"Parameter Table\"") != std::wstring::npos
			   && x.find (L"</Tab></Tabs></Ribbon>") != std::wstring::npos, "a Ribbon / Tabs / Tab document");
	Check (x.find (L"<Group FTCommand=\"_ParamTableAddIn_.ParamTableGroup\"") != std::wstring::npos,
		   "the group names its hidden function as APPLICATION.FUNCTION");
	Check (x.find (L"<Button FTCommand=\"_ParamTableAddIn_.LatheParamsDumpEntry\" KeyTip=\"D\" LargeMode=\"Always\" />")
			   != std::wstring::npos && Count (x, L"<Button ") == 2, "each button names its function the same way, large");
	for (const wchar_t *tag : { L"Ribbon", L"Tabs", L"Tab ", L"Groups", L"Group ", L"Elements" })
		{
		const std::wstring t (tag), name = t.back () == L' ' ? t.substr (0, t.size () - 1) : t;
		if (Count (x, L"<" + t) != Count (x, L"</" + name + L">"))
			Check (false, "every tag closed");
		}

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
