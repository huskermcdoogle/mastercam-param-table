#include "Csv.h"

#include <cmath>
#include <cstdio>
#include <cstdlib>
#include <cwchar>
#include <cwctype>

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>

namespace Csv
	{
	namespace
		{
		bool NeedsQuote (const std::wstring &s)
			{
			for (wchar_t c : s)
				if (c == L',' || c == L'"' || c == L'\r' || c == L'\n')
					return true;
			// Leading or trailing space would be eaten by a careless reader.
			return !s.empty () && (s.front () == L' ' || s.back () == L' ');
			}

		std::wstring TrimImpl (const std::wstring &s)
			{
			const size_t a = s.find_first_not_of (L" \t\r\n");
			if (a == std::wstring::npos)
				return L"";
			const size_t b = s.find_last_not_of (L" \t\r\n");
			return s.substr (a, b - a + 1);
			}
		}

	std::wstring Serialize (const std::vector<Row> &rows)
		{
		std::wstring out;
		for (const Row &row : rows)
			{
			for (size_t i = 0; i < row.size (); ++i)
				{
				if (i != 0)
					out += L',';
				if (NeedsQuote (row[i]))
					{
					out += L'"';
					for (wchar_t c : row[i])
						{
						if (c == L'"')
							out += L'"';
						out += c;
						}
					out += L'"';
					}
				else
					out += row[i];
				}
			out += L"\r\n";
			}
		return out;
		}

	std::vector<Row> Parse (const std::wstring &input)
		{
		std::vector<Row> rows;
		std::wstring text = input;
		if (!text.empty () && text[0] == 0xFEFF)
			text.erase (0, 1);

		Row row;
		std::wstring cell;
		bool inQuotes = false;
		bool cellWasQuoted = false;

		auto endCell = [&] ()
			{
			row.push_back (cell);
			cell.clear ();
			cellWasQuoted = false;
			};

		auto endRow = [&] ()
			{
			endCell ();
			// A fully blank line - one empty unquoted cell - is not a row.
			const bool blank = (row.size () == 1 && row[0].empty ());
			if (!blank)
				rows.push_back (row);
			row.clear ();
			};

		for (size_t i = 0; i < text.size (); ++i)
			{
			const wchar_t c = text[i];

			if (inQuotes)
				{
				if (c == L'"')
					{
					if (i + 1 < text.size () && text[i + 1] == L'"')
						{
						cell += L'"';
						++i;
						}
					else
						inQuotes = false;
					}
				else
					cell += c;
				continue;
				}

			if (c == L'"' && cell.empty () && !cellWasQuoted)
				{
				inQuotes = true;
				cellWasQuoted = true;
				}
			else if (c == L',')
				endCell ();
			else if (c == L'\r')
				{
				if (i + 1 < text.size () && text[i + 1] == L'\n')
					++i;
				endRow ();
				}
			else if (c == L'\n')
				endRow ();
			else
				cell += c;
			}

		// No trailing newline: the last row is still a row.
		if (!cell.empty () || !row.empty () || cellWasQuoted)
			endRow ();

		return rows;
		}

	std::string ToUtf8Bom (const std::wstring &text)
		{
		std::string out = "\xEF\xBB\xBF";
		if (text.empty ())
			return out;

		const int n = ::WideCharToMultiByte (CP_UTF8, 0, text.data (),
											 static_cast<int> (text.size ()),
											 nullptr, 0, nullptr, nullptr);
		std::string body (static_cast<size_t> (n), '\0');
		::WideCharToMultiByte (CP_UTF8, 0, text.data (),
							   static_cast<int> (text.size ()),
							   &body[0], n, nullptr, nullptr);
		return out + body;
		}

	std::wstring FromUtf8 (const std::string &bytes)
		{
		std::string b = bytes;
		if (b.size () >= 3 && static_cast<unsigned char> (b[0]) == 0xEF
			&& static_cast<unsigned char> (b[1]) == 0xBB
			&& static_cast<unsigned char> (b[2]) == 0xBF)
			b.erase (0, 3);
		if (b.empty ())
			return L"";

		// Not valid UTF-8 means it was saved as ANSI - which is what Excel does
		// to a "CSV" that is not the UTF-8 flavour. Fall back rather than fail.
		UINT cp = CP_UTF8;
		int n = ::MultiByteToWideChar (cp, MB_ERR_INVALID_CHARS, b.data (),
									   static_cast<int> (b.size ()), nullptr, 0);
		if (n == 0)
			{
			cp = CP_ACP;
			n = ::MultiByteToWideChar (cp, 0, b.data (),
									   static_cast<int> (b.size ()), nullptr, 0);
			}
		std::wstring out (static_cast<size_t> (n), L'\0');
		::MultiByteToWideChar (cp, 0, b.data (), static_cast<int> (b.size ()),
							   &out[0], n);
		return out;
		}

	std::wstring Trim (const std::wstring &s)
		{
		return TrimImpl (s);
		}

	std::wstring FormatDouble (double v)
		{
		if (v == 0.0)
			return L"0";			// also swallows negative zero

		wchar_t buf[40];
		// Shortest first: 15 significant digits round-trips almost everything a
		// person types; 17 always does. Take the first that gets the same
		// double back.
		for (int digits : { 15, 16, 17 })
			{
			swprintf_s (buf, L"%.*g", digits, v);
			if (wcstod (buf, nullptr) == v)
				break;
			}
		return buf;
		}

	std::wstring Tidy (double v)
		{
		if (v == 0.0 || !std::isfinite (v))
			return FormatDouble (v);

		// THE SHORTEST DECIMAL THAT IS STILL THIS VALUE, as a load judges it: within
		// half of SameDouble's tolerance, so a dump loaded back untouched can never
		// read as a change. Mastercam's noise - 0.001 stored through a float as
		// 0.0010000000474975, 0.07 as 0.070000000000001 - is far smaller and goes.
		const double tol = 0.5e-9 * (std::max) (1.0, std::fabs (v));
		const int mag = static_cast<int> (std::floor (std::log10 (std::fabs (v))));
		wchar_t buf[96];
		std::wstring s;
		for (int sig = 1; sig <= 17; ++sig)
			{
			int dec = sig - 1 - mag;
			if (dec < 0)
				dec = 0;
			if (dec > 40)
				dec = 40;
			swprintf_s (buf, L"%.*f", dec, v);
			if (std::fabs (wcstod (buf, nullptr) - v) <= tol)
				break;
			}
		s = buf;
		if (s.find (L'.') != std::wstring::npos)
			{
			while (!s.empty () && s.back () == L'0')
				s.pop_back ();
			if (!s.empty () && s.back () == L'.')
				s.pop_back ();
			}
		if (s == L"-0")
			s = L"0";
		return s;
		}

	bool ParseDouble (const std::wstring &text, double &out)
		{
		const std::wstring t = TrimImpl (text);
		if (t.empty ())
			return false;

		wchar_t *end = nullptr;
		const double v = wcstod (t.c_str (), &end);
		if (end == t.c_str () || *end != L'\0')
			return false;
		if (!std::isfinite (v))
			return false;

		out = v;
		return true;
		}

	bool ParseLong (const std::wstring &text, long long &out)
		{
		double d = 0;
		if (!ParseDouble (text, d))
			return false;
		if (d != std::floor (d) || std::fabs (d) > 2147483647.0)
			return false;
		out = static_cast<long long> (d);
		return true;
		}

	bool ParseBool (const std::wstring &text, bool &out)
		{
		std::wstring t = TrimImpl (text);
		for (wchar_t &c : t)
			c = static_cast<wchar_t> (std::towlower (c));

		if (t == L"1" || t == L"true" || t == L"yes" || t == L"on")
			{
			out = true;
			return true;
			}
		if (t == L"0" || t == L"false" || t == L"no" || t == L"off")
			{
			out = false;
			return true;
			}
		return false;
		}

	bool SameDouble (double a, double b)
		{
		if (a == b)
			return true;
		const double scale = (std::max) (std::fabs (a), std::fabs (b));
		return std::fabs (a - b) <= 1e-9 * (std::max) (1.0, scale);
		}
	}
