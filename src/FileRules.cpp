#include "FileRules.h"

#include <cwchar>

namespace FileRules
	{
	std::wstring Name (const std::wstring &pattern, const std::wstring &part, std::time_t when,
					   bool selectedOnly, size_t ops, const std::wstring &ext)
		{
		std::tm tmv = {};
		localtime_s (&tmv, &when);
		wchar_t date[16] = L"", time[16] = L"";
		std::wcsftime (date, 16, L"%Y%m%d", &tmv);
		std::wcsftime (time, 16, L"%H%M%S", &tmv);

		std::wstring p = pattern.find_first_not_of (L" \t") == std::wstring::npos ? kDefaultPattern : pattern;
		std::wstring out;
		for (size_t i = 0; i < p.size ();)
			{
			if (p[i] == L'{')
				{
				const size_t close = p.find (L'}', i);
				if (close != std::wstring::npos)
					{
					const std::wstring tok = p.substr (i + 1, close - i - 1);
					std::wstring v;
					bool known = true;
					if (tok == L"part") v = part;
					else if (tok == L"date") v = date;
					else if (tok == L"time") v = time;
					else if (tok == L"scope") v = selectedOnly ? L"selected" : L"all";
					else if (tok == L"ops") v = std::to_wstring (ops);
					else known = false;
					if (known)
						{
						out += v;
						i = close + 1;
						continue;
						}
					}
				}
			out += p[i++];
			}

		// What Windows refuses in a file name.
		for (wchar_t &c : out)
			if (c < 32 || std::wcschr (L"<>:\"/\\|?*", c) != nullptr)
				c = L'_';
		while (!out.empty () && (out.back () == L' ' || out.back () == L'.'))
			out.pop_back ();
		if (out.empty ())
			out = L"lathe_params";

		// A typed extension of either kind gives way to the one chosen.
		std::wstring lower = out;
		for (wchar_t &c : lower)
			c = static_cast<wchar_t> (towlower (c));
		if (lower.size () >= 5 && (lower.compare (lower.size () - 5, 5, L".xlsx") == 0
								   || lower.compare (lower.size () - 5, 5, L".xlsm") == 0))
			out.resize (out.size () - 5);
		return out + ext;
		}

	std::filesystem::path Unique (const std::filesystem::path &folder, const std::wstring &name)
		{
		std::filesystem::path first = folder / name;
		std::error_code ec;
		if (!std::filesystem::exists (first, ec))
			return first;
		const std::wstring stem = first.stem ().wstring ();
		const std::wstring ext = first.extension ().wstring ();
		for (int n = 2; n < 10000; ++n)
			{
			std::filesystem::path p = folder / (stem + L" (" + std::to_wstring (n) + L")" + ext);
			if (!std::filesystem::exists (p, ec))
				return p;
			}
		return first;
		}
	}
