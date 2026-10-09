#include "stdafx.h"
#include "Settings.h"
#include "FileRules.h"

namespace
	{
	const wchar_t *const kKey = L"Software\\ParamTableTool";

	std::wstring GetText (const wchar_t *name, const std::wstring &dflt)
		{
		wchar_t buf[2048] = L"";
		DWORD size = sizeof (buf);
		if (RegGetValueW (HKEY_CURRENT_USER, kKey, name, RRF_RT_REG_SZ, nullptr, buf, &size) != ERROR_SUCCESS)
			return dflt;
		return buf;
		}

	bool GetFlag (const wchar_t *name, bool dflt)
		{
		DWORD v = 0, size = sizeof (v);
		if (RegGetValueW (HKEY_CURRENT_USER, kKey, name, RRF_RT_REG_DWORD, nullptr, &v, &size) != ERROR_SUCCESS)
			return dflt;
		return v != 0;
		}

	void SetText (const wchar_t *name, const std::wstring &v)
		{
		RegSetKeyValueW (HKEY_CURRENT_USER, kKey, name, REG_SZ, v.c_str (),
						 static_cast<DWORD> ((v.size () + 1) * sizeof (wchar_t)));
		}

	void SetFlag (const wchar_t *name, bool v)
		{
		const DWORD d = v ? 1 : 0;
		RegSetKeyValueW (HKEY_CURRENT_USER, kKey, name, REG_DWORD, &d, sizeof (d));
		}
	}

namespace Settings
	{
	Dump LoadDump ()
		{
		Dump d;
		d.folder = GetText (L"DumpFolder", L"");
		d.pattern = GetText (L"DumpPattern", FileRules::kDefaultPattern);
		d.openExcel = GetFlag (L"OpenExcel", true);
		d.pictures = GetFlag (L"ToolPictures", true);
		d.macros = GetFlag (L"Macros", false);
		d.skipKinds = GetText (L"SkipKinds", L"");
		return d;
		}

	void SaveDump (const Dump &d)
		{
		SetText (L"DumpFolder", d.folder);
		SetText (L"DumpPattern", d.pattern);
		SetFlag (L"OpenExcel", d.openExcel);
		SetFlag (L"ToolPictures", d.pictures);
		SetFlag (L"Macros", d.macros);
		SetText (L"SkipKinds", d.skipKinds);
		}

	void SetLastDump (const std::wstring &file, const std::wstring &part)
		{
		SetText (L"LastDumpFile", file);
		SetText (L"LastDumpPart", part);
		}

	unsigned long Diag ()
		{
		// A file beside the DLL first (Add-Ins\ParamTable\diag.txt holding a number) -
		// for machines where the registry cannot be edited by hand.
		HMODULE self = nullptr;
		wchar_t path[MAX_PATH] = {};
		if (GetModuleHandleExW (GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
								reinterpret_cast<LPCWSTR> (&Diag), &self)
			&& GetModuleFileNameW (self, path, MAX_PATH) > 0)
			{
			std::wstring file (path);
			file = file.substr (0, file.find_last_of (L"\\/") + 1) + L"diag.txt";
			FILE *f = nullptr;
			if (_wfopen_s (&f, file.c_str (), L"r") == 0 && f != nullptr)
				{
				unsigned long v = 0;
				const int got = fscanf_s (f, "%lu", &v);
				fclose (f);
				if (got == 1)
					return v;
				}
			}
		DWORD v = 0, size = sizeof (v);
		if (RegGetValueW (HKEY_CURRENT_USER, kKey, L"Diag", RRF_RT_REG_DWORD, nullptr, &v, &size) != ERROR_SUCCESS)
			return 0;
		return v;
		}

	std::wstring LastDumpFor (const std::wstring &part)
		{
		return _wcsicmp (GetText (L"LastDumpPart", L"").c_str (), part.c_str ()) == 0
				   ? GetText (L"LastDumpFile", L"") : std::wstring ();
		}
	}
