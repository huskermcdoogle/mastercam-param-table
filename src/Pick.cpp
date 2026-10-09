#include "Pick.h"

#include <cwctype>

namespace
	{
	std::wstring Lower (const std::wstring &s)
		{
		std::wstring o;
		for (wchar_t c : s)
			o += static_cast<wchar_t> (std::towlower (c));
		return o;
		}
	}

namespace Pick
	{
	bool Matches (const std::wstring &line, const std::wstring &typed)
		{
		const std::wstring hay = Lower (line), q = Lower (typed);
		bool any = false;
		for (size_t i = 0; i < q.size ();)
			{
			while (i < q.size () && std::iswspace (q[i]))
				++i;
			size_t e = i;
			while (e < q.size () && !std::iswspace (q[e]))
				++e;
			if (e > i)
				{
				if (hay.find (q.substr (i, e - i)) == std::wstring::npos)
					return false;
				any = true;
				}
			i = e;
			}
		return any;
		}

	std::wstring IdsText (const std::vector<long> &ids)
		{
		std::wstring o;
		for (long id : ids)
			o += (o.empty () ? L"" : L",") + std::to_wstring (id);
		return o;
		}

	std::vector<long> ParseIds (const std::wstring &text)
		{
		std::vector<long> ids;
		for (size_t i = 0; i < text.size ();)
			{
			while (i < text.size () && !(text[i] >= L'0' && text[i] <= L'9') && text[i] != L'-')
				++i;
			size_t e = i;
			if (e < text.size () && text[e] == L'-')
				++e;
			const size_t d0 = e;
			while (e < text.size () && text[e] >= L'0' && text[e] <= L'9' && e - d0 < 9)
				++e;
			if (e > d0)
				ids.push_back (std::stol (text.substr (i, e - i)));
			i = e > i ? e : i + 1;
			}
		return ids;
		}

	std::wstring PartKey (const std::wstring &partFile)
		{
		const size_t slash = partFile.find_last_of (L"\\/");
		std::wstring name = Lower (slash == std::wstring::npos ? partFile : partFile.substr (slash + 1));
		// A registry key name: no backslash (cannot happen in a file name), and
		// Windows keeps 255 characters.
		if (name.size () > 200)
			name = name.substr (0, 200);
		return name;
		}

	std::wstring CleanName (const std::wstring &name)
		{
		size_t a = 0, b = name.size ();
		while (a < b && std::iswspace (name[a]))
			++a;
		while (b > a && std::iswspace (name[b - 1]))
			--b;
		std::wstring o;
		for (size_t i = a; i < b; ++i)
			o += name[i] < 0x20 ? L' ' : name[i];
		return o.size () > 60 ? o.substr (0, 60) : o;
		}
	}
