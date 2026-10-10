#include "stdafx.h"
#include "Settings.h"
#include "FileRules.h"
#include "Pick.h"

#include <algorithm>

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
		d.stockSim = GetFlag (L"StockSim", true);
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
		SetFlag (L"StockSim", d.stockSim);
		SetText (L"SkipKinds", d.skipKinds);
		}

	void SetLastDump (const std::wstring &file, const std::wstring &part)
		{
		SetText (L"LastDumpFile", file);
		SetText (L"LastDumpPart", part);
		}

	std::wstring AddinFolder ()
		{
		HMODULE self = nullptr;
		wchar_t path[MAX_PATH] = {};
		if (GetModuleHandleExW (GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
								reinterpret_cast<LPCWSTR> (&AddinFolder), &self)
			&& GetModuleFileNameW (self, path, MAX_PATH) > 0)
			{
			const std::wstring file (path);
			return file.substr (0, file.find_last_of (L"\\/") + 1);
			}
		return std::wstring ();
		}

	unsigned long Diag ()
		{
		// A file beside the DLL first (Add-Ins\ParamTable\diag.txt holding a number) -
		// for machines where the registry cannot be edited by hand.
		const std::wstring folder = AddinFolder ();
		if (!folder.empty ())
			{
			const std::wstring file = folder + L"diag.txt";
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

	namespace
		{
		std::wstring SelectionsKey (const std::wstring &partFile)
			{
			return std::wstring (kKey) + L"\\Selections\\" + Pick::PartKey (partFile);
			}
		}

	std::vector<std::pair<std::wstring, std::vector<long>>> Selections (const std::wstring &partFile)
		{
		std::vector<std::pair<std::wstring, std::vector<long>>> out;
		HKEY key = nullptr;
		if (RegOpenKeyExW (HKEY_CURRENT_USER, SelectionsKey (partFile).c_str (), 0, KEY_READ, &key) != ERROR_SUCCESS)
			return out;
		for (DWORD i = 0;; ++i)
			{
			wchar_t name[256] = L"";
			wchar_t data[4096] = L"";
			DWORD nameLen = 256, type = 0, size = sizeof (data) - sizeof (wchar_t);
			const LSTATUS r = RegEnumValueW (key, i, name, &nameLen, nullptr, &type,
											 reinterpret_cast<BYTE *> (data), &size);
			if (r == ERROR_NO_MORE_ITEMS)
				break;
			if (r != ERROR_SUCCESS || type != REG_SZ)
				continue;				// too long, or not ours: skipped, never fatal
			data[size / sizeof (wchar_t)] = 0;
			out.push_back ({ name, Pick::ParseIds (data) });
			}
		RegCloseKey (key);
		std::sort (out.begin (), out.end (), [] (const auto &a, const auto &b)
			{ return _wcsicmp (a.first.c_str (), b.first.c_str ()) < 0; });
		return out;
		}

	void SaveSelection (const std::wstring &partFile, const std::wstring &name, const std::vector<long> &ids)
		{
		const std::wstring v = Pick::IdsText (ids);
		RegSetKeyValueW (HKEY_CURRENT_USER, SelectionsKey (partFile).c_str (), name.c_str (), REG_SZ, v.c_str (),
						 static_cast<DWORD> ((v.size () + 1) * sizeof (wchar_t)));
		}

	void DeleteSelection (const std::wstring &partFile, const std::wstring &name)
		{
		RegDeleteKeyValueW (HKEY_CURRENT_USER, SelectionsKey (partFile).c_str (), name.c_str ());
		}

	std::wstring LastDumpFor (const std::wstring &part)
		{
		return _wcsicmp (GetText (L"LastDumpPart", L"").c_str (), part.c_str ()) == 0
				   ? GetText (L"LastDumpFile", L"") : std::wstring ();
		}
	}
