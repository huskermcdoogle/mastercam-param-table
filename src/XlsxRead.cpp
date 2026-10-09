//
// XlsxRead.cpp - reading back a workbook Excel saved. No Mastercam SDK.
//
// Three layers, each small enough to check by eye:
//
//   Inflate    - RFC 1951 DEFLATE. Excel compresses every part of a saved
//                workbook; this is the whole of what that takes. Canonical
//                Huffman decoding after the reference decoder (zlib's puff).
//   Unzip      - the zip's central directory, each part decompressed and its
//                CRC checked: a damaged file is refused, not half-read.
//   ReadSheet  - the first sheet's cells as TEXT, as a CSV of the same sheet
//                would hold them, so every rule the load applies is unchanged.
//
#include "Xlsx.h"
#include "Csv.h"

#include <algorithm>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <iterator>
#include <cwctype>

namespace
	{
	// ---- DEFLATE ----------------------------------------------------------

	struct BitIn
		{
		const unsigned char *p = nullptr;
		size_t n = 0;
		size_t i = 0;
		uint32_t buf = 0;
		int cnt = 0;
		bool bad = false;

		int Bits (int need)
			{
			uint32_t v = buf;
			while (cnt < need)
				{
				if (i >= n)
					{
					bad = true;
					return 0;
					}
				v |= static_cast<uint32_t> (p[i++]) << cnt;
				cnt += 8;
				}
			buf = v >> need;
			cnt -= need;
			return static_cast<int> (v & ((1u << need) - 1u));
			}
		};

	struct Huffman
		{
		short count[16] = {};		//!< codes of each length
		short symbol[320] = {};		//!< symbols ordered by code
		};

	/// Canonical codes from code lengths. < 0 = over-subscribed (damaged data).
	int Build (Huffman &h, const short *length, int n)
		{
		std::memset (h.count, 0, sizeof (h.count));
		for (int s = 0; s < n; ++s)
			++h.count[length[s]];
		if (h.count[0] == n)
			return 0;
		int left = 1;
		for (int len = 1; len < 16; ++len)
			{
			left <<= 1;
			left -= h.count[len];
			if (left < 0)
				return left;
			}
		short offs[16];
		offs[1] = 0;
		for (int len = 1; len < 15; ++len)
			offs[len + 1] = static_cast<short> (offs[len] + h.count[len]);
		for (int s = 0; s < n; ++s)
			if (length[s] != 0)
				h.symbol[offs[length[s]]++] = static_cast<short> (s);
		return left;
		}

	int Decode (BitIn &in, const Huffman &h)
		{
		int code = 0, first = 0, index = 0;
		for (int len = 1; len < 16; ++len)
			{
			code |= in.Bits (1);
			if (in.bad)
				return -1;
			const int count = h.count[len];
			if (code - count < first)
				return h.symbol[index + (code - first)];
			index += count;
			first += count;
			first <<= 1;
			code <<= 1;
			}
		return -1;
		}

	const short kLenBase[29] = { 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
								 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258 };
	const short kLenExtra[29] = { 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2,
								  3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0 };
	const short kDistBase[30] = { 1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
								  257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145,
								  8193, 12289, 16385, 24577 };
	const short kDistExtra[30] = { 0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6,
								   7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13 };

	bool Codes (BitIn &in, std::string &out, const Huffman &lens, const Huffman &dists)
		{
		for (;;)
			{
			int sym = Decode (in, lens);
			if (sym < 0)
				return false;
			if (sym < 256)
				{
				out += static_cast<char> (sym);
				continue;
				}
			if (sym == 256)
				return true;
			sym -= 257;
			if (sym >= 29)
				return false;
			const int len = kLenBase[sym] + in.Bits (kLenExtra[sym]);
			const int dsym = Decode (in, dists);
			if (dsym < 0 || dsym >= 30)
				return false;
			const size_t dist = static_cast<size_t> (kDistBase[dsym] + in.Bits (kDistExtra[dsym]));
			if (in.bad || dist > out.size ())
				return false;
			const size_t from = out.size () - dist;
			for (int k = 0; k < len; ++k)
				out += out[from + static_cast<size_t> (k)];
			}
		}

	bool Fixed (BitIn &in, std::string &out)
		{
		static Huffman lens, dists;
		static bool built = false;
		if (!built)
			{
			short l[288];
			int s = 0;
			for (; s < 144; ++s) l[s] = 8;
			for (; s < 256; ++s) l[s] = 9;
			for (; s < 280; ++s) l[s] = 7;
			for (; s < 288; ++s) l[s] = 8;
			Build (lens, l, 288);
			for (s = 0; s < 30; ++s) l[s] = 5;
			Build (dists, l, 30);
			built = true;
			}
		return Codes (in, out, lens, dists);
		}

	bool Dynamic (BitIn &in, std::string &out)
		{
		static const short order[19] = { 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15 };
		const int nlen = in.Bits (5) + 257;
		const int ndist = in.Bits (5) + 1;
		const int ncode = in.Bits (4) + 4;
		if (in.bad || nlen > 286 || ndist > 30)
			return false;

		short lengths[320] = {};
		for (int k = 0; k < ncode; ++k)
			lengths[order[k]] = static_cast<short> (in.Bits (3));
		Huffman code;
		if (Build (code, lengths, 19) != 0)
			return false;

		int at = 0;
		while (at < nlen + ndist)
			{
			int sym = Decode (in, code);
			if (sym < 0)
				return false;
			if (sym < 16)
				{
				lengths[at++] = static_cast<short> (sym);
				continue;
				}
			short len = 0;
			int repeat = 0;
			if (sym == 16)
				{
				if (at == 0)
					return false;
				len = lengths[at - 1];
				repeat = 3 + in.Bits (2);
				}
			else if (sym == 17)
				repeat = 3 + in.Bits (3);
			else
				repeat = 11 + in.Bits (7);
			if (in.bad || at + repeat > nlen + ndist)
				return false;
			while (repeat--)
				lengths[at++] = len;
			}
		if (lengths[256] == 0)
			return false;

		Huffman lens, dists;
		if (Build (lens, lengths, nlen) < 0 || Build (dists, lengths + nlen, ndist) < 0)
			return false;
		return Codes (in, out, lens, dists);
		}

	// ---- Zip ----------------------------------------------------------------

	uint32_t Crc32 (const std::string &d)
		{
		static uint32_t table[256];
		static bool init = false;
		if (!init)
			{
			for (uint32_t n = 0; n < 256; ++n)
				{
				uint32_t c = n;
				for (int k = 0; k < 8; ++k)
					c = (c & 1) ? 0xEDB88320u ^ (c >> 1) : c >> 1;
				table[n] = c;
				}
			init = true;
			}
		uint32_t c = 0xFFFFFFFFu;
		for (unsigned char b : d)
			c = table[(c ^ b) & 0xFF] ^ (c >> 8);
		return c ^ 0xFFFFFFFFu;
		}

	uint32_t U16 (const std::string &s, size_t at)
		{
		return static_cast<unsigned char> (s[at]) | (static_cast<unsigned char> (s[at + 1]) << 8);
		}
	uint32_t U32 (const std::string &s, size_t at)
		{
		return U16 (s, at) | (U16 (s, at + 2) << 16);
		}

	// ---- XML, just enough ----------------------------------------------------

	/// XML text to plain text: the five entities, numeric references, and
	/// Excel's own _xHHHH_ escape (how it keeps a carriage return in a cell).
	std::wstring Unescape (const std::string &utf8)
		{
		const std::wstring w = Csv::FromUtf8 (utf8);
		std::wstring o;
		for (size_t i = 0; i < w.size (); ++i)
			{
			if (w[i] == L'&')
				{
				const size_t semi = w.find (L';', i);
				if (semi != std::wstring::npos && semi - i <= 10)
					{
					const std::wstring e = w.substr (i + 1, semi - i - 1);
					wchar_t c = 0;
					if (e == L"amp") c = L'&';
					else if (e == L"lt") c = L'<';
					else if (e == L"gt") c = L'>';
					else if (e == L"quot") c = L'"';
					else if (e == L"apos") c = L'\'';
					else if (e.size () > 1 && e[0] == L'#')
						c = static_cast<wchar_t> (e[1] == L'x' ? wcstoul (e.c_str () + 2, nullptr, 16)
															   : wcstoul (e.c_str () + 1, nullptr, 10));
					if (c != 0)
						{
						o += c;
						i = semi;
						continue;
						}
					}
				}
			if (w[i] == L'_' && i + 6 < w.size () && w[i + 1] == L'x' && w[i + 6] == L'_')
				{
				bool hex = true;
				for (size_t k = i + 2; k < i + 6; ++k)
					hex = hex && iswxdigit (w[k]);
				if (hex)
					{
					o += static_cast<wchar_t> (wcstoul (w.substr (i + 2, 4).c_str (), nullptr, 16));
					i += 6;
					continue;
					}
				}
			o += w[i];
			}
		return o;
		}

	/// The next element named `name` at or after `from`: [start of "<name", end
	/// of its open tag]. A name must be followed by a space, '>' or '/', so "c"
	/// does not match "col".
	bool FindTag (const std::string &x, const std::string &name, size_t from,
				  size_t &start, size_t &openEnd)
		{
		for (size_t at = x.find ("<" + name, from); at != std::string::npos;
			 at = x.find ("<" + name, at + 1))
			{
			const size_t after = at + 1 + name.size ();
			if (after < x.size () && (x[after] == ' ' || x[after] == '>' || x[after] == '/'))
				{
				const size_t gt = x.find ('>', after);
				if (gt == std::string::npos)
					return false;
				start = at;
				openEnd = gt + 1;
				return true;
				}
			}
		return false;
		}

	bool SelfClosing (const std::string &x, size_t openEnd)
		{
		return openEnd >= 2 && x[openEnd - 2] == '/';
		}

	/// An attribute of the open tag [start, openEnd), or "".
	std::string Attr (const std::string &x, size_t start, size_t openEnd, const std::string &name)
		{
		const std::string key = " " + name + "=\"";
		const size_t at = x.find (key, start);
		if (at == std::string::npos || at >= openEnd)
			return std::string ();
		const size_t v = at + key.size ();
		const size_t q = x.find ('"', v);
		return q == std::string::npos ? std::string () : x.substr (v, q - v);
		}

	/// The content of an element whose open tag ends at openEnd.
	std::string Body (const std::string &x, const std::string &name, size_t openEnd, size_t &end)
		{
		const size_t close = x.find ("</" + name + ">", openEnd);
		end = close == std::string::npos ? x.size () : close + name.size () + 3;
		return x.substr (openEnd, (close == std::string::npos ? x.size () : close) - openEnd);
		}

	/// All the <t> text inside a fragment, joined - a rich-text cell is split
	/// into runs, and phonetic guides (<rPh>) are not part of the value.
	std::wstring TextOf (std::string frag)
		{
		for (size_t s, e; FindTag (frag, "rPh", 0, s, e);)
			{
			size_t end = 0;
			if (!SelfClosing (frag, e))
				Body (frag, "rPh", e, end);
			else
				end = e;
			frag.erase (s, end - s);
			}
		std::wstring out;
		size_t from = 0, s = 0, e = 0;
		while (FindTag (frag, "t", from, s, e))
			{
			if (SelfClosing (frag, e))
				{
				from = e;
				continue;
				}
			size_t end = 0;
			out += Unescape (Body (frag, "t", e, end));
			from = end;
			}
		return out;
		}

	/// "AB12" -> column 27 (0-based), row 12. False when there is no letter.
	bool CellRef (const std::string &ref, size_t &col, size_t &row)
		{
		size_t i = 0, c = 0;
		while (i < ref.size () && ref[i] >= 'A' && ref[i] <= 'Z')
			c = c * 26 + static_cast<size_t> (ref[i++] - 'A' + 1);
		if (i == 0)
			return false;
		col = c - 1;
		row = static_cast<size_t> (std::strtoul (ref.c_str () + i, nullptr, 10));
		return true;
		}

	/// The worksheet part the workbook lists first (or the one named "Lathe
	/// params"), through the workbook's relationships. With `only`, the sheet
	/// of exactly that name or "" - another sheet is never a stand-in for it.
	std::string SheetPart (const std::map<std::string, std::string> &parts,
						   const std::string &only = std::string ())
		{
		const auto wb = parts.find ("xl/workbook.xml");
		const auto rels = parts.find ("xl/_rels/workbook.xml.rels");
		if (wb == parts.end () || rels == parts.end ())
			return only.empty () ? "xl/worksheets/sheet1.xml" : std::string ();

		std::string rid;
		size_t from = 0, s = 0, e = 0;
		while (FindTag (wb->second, "sheet", from, s, e))
			{
			const std::string id = Attr (wb->second, s, e, "r:id");
			const std::string name = Attr (wb->second, s, e, "name");
			if (only.empty () ? (rid.empty () || name == "Lathe params") : name == only)
				rid = id;
			from = e;
			}
		if (!only.empty () && rid.empty ())
			return std::string ();
		from = 0;
		while (FindTag (rels->second, "Relationship", from, s, e))
			{
			if (Attr (rels->second, s, e, "Id") == rid)
				{
				std::string target = Attr (rels->second, s, e, "Target");
				if (!target.empty () && target[0] == '/')
					return target.substr (1);
				return "xl/" + target;
				}
			from = e;
			}
		return "xl/worksheets/sheet1.xml";
		}
	}

namespace Xlsx
	{
	bool Inflate (const std::string &in, std::string &out)
		{
		BitIn b;
		b.p = reinterpret_cast<const unsigned char *> (in.data ());
		b.n = in.size ();
		out.clear ();
		for (;;)
			{
			const int last = b.Bits (1);
			const int type = b.Bits (2);
			if (b.bad)
				return false;
			bool ok = false;
			if (type == 0)
				{
				// Stored: whole bytes, after dropping the bits of the current one.
				b.buf = 0;
				b.cnt = 0;
				if (b.i + 4 > b.n)
					return false;
				const uint32_t len = b.p[b.i] | (b.p[b.i + 1] << 8);
				const uint32_t nlen = b.p[b.i + 2] | (b.p[b.i + 3] << 8);
				b.i += 4;
				if (len != (~nlen & 0xFFFFu) || b.i + len > b.n)
					return false;
				out.append (reinterpret_cast<const char *> (b.p + b.i), len);
				b.i += len;
				ok = true;
				}
			else if (type == 1)
				ok = Fixed (b, out);
			else if (type == 2)
				ok = Dynamic (b, out);
			if (!ok || b.bad)
				return false;
			if (last)
				return true;
			}
		}

	bool Unzip (const std::string &zip, std::map<std::string, std::string> &parts,
				std::wstring &why)
		{
		parts.clear ();
		// The end-of-central-directory record, searched from the end (a zip
		// comment can follow it).
		size_t eocd = std::string::npos;
		if (zip.size () >= 22)
			for (size_t at = zip.size () - 22;; --at)
				{
				if (U32 (zip, at) == 0x06054B50u)
					{
					eocd = at;
					break;
					}
				if (at == 0 || zip.size () - at > 22 + 65535)
					break;
				}
		if (eocd == std::string::npos)
			{
			why = L"it is not a workbook (no zip directory)";
			return false;
			}

		const uint32_t count = U16 (zip, eocd + 10);
		size_t at = U32 (zip, eocd + 16);
		for (uint32_t k = 0; k < count; ++k)
			{
			if (at + 46 > zip.size () || U32 (zip, at) != 0x02014B50u)
				{
				why = L"the workbook's zip directory is damaged";
				return false;
				}
			const uint32_t method = U16 (zip, at + 10);
			const uint32_t crc = U32 (zip, at + 16);
			const uint32_t csize = U32 (zip, at + 20);
			const uint32_t usize = U32 (zip, at + 24);
			const uint32_t nlen = U16 (zip, at + 28);
			const uint32_t elen = U16 (zip, at + 30);
			const uint32_t clen = U16 (zip, at + 32);
			const uint32_t local = U32 (zip, at + 42);
			const std::string name = zip.substr (at + 46, nlen);
			at += 46 + nlen + elen + clen;

			if (local + 30 > zip.size () || U32 (zip, local) != 0x04034B50u)
				{
				why = L"part " + Csv::FromUtf8 (name) + L" is damaged";
				return false;
				}
			const size_t data = local + 30 + U16 (zip, local + 26) + U16 (zip, local + 28);
			if (data + csize > zip.size ())
				{
				why = L"part " + Csv::FromUtf8 (name) + L" is cut short";
				return false;
				}
			const std::string packed = zip.substr (data, csize);
			std::string bytes;
			if (method == 0)
				bytes = packed;
			else if (method == 8)
				{
				if (!Inflate (packed, bytes))
					{
					why = L"part " + Csv::FromUtf8 (name) + L" could not be decompressed";
					return false;
					}
				}
			else
				{
				why = L"part " + Csv::FromUtf8 (name) + L" uses an unsupported compression ("
					  + std::to_wstring (method) + L")";
				return false;
				}
			if (bytes.size () != usize || Crc32 (bytes) != crc)
				{
				why = L"part " + Csv::FromUtf8 (name) + L" fails its checksum - the file is damaged";
				return false;
				}
			parts[name] = bytes;
			}
		return true;
		}

	bool ReadSheetBytes (const std::string &bytes,
						 std::vector<std::vector<std::wstring>> &rows,
						 std::vector<size_t> &sheetRow, std::wstring &why)
		{
		return ReadNamedSheetBytes (bytes, std::wstring (), rows, sheetRow, why);
		}

	bool ReadGridBytes (const std::string &bytes, const std::wstring &sheetName, Grid &grid, std::wstring &why)
		{
		grid.clear ();
		std::map<std::string, std::string> parts;
		if (!Unzip (bytes, parts, why))
			return false;

		// Excel moves every string it saves into one table.
		std::vector<std::wstring> shared;
		const auto ss = parts.find ("xl/sharedStrings.xml");
		if (ss != parts.end ())
			{
			size_t from = 0, s = 0, e = 0;
			while (FindTag (ss->second, "si", from, s, e))
				{
				if (SelfClosing (ss->second, e))
					{
					shared.push_back (std::wstring ());
					from = e;
					continue;
					}
				size_t end = 0;
				shared.push_back (TextOf (Body (ss->second, "si", e, end)));
				from = end;
				}
			}

		// The workbook lists sheet names as UTF-8; ToUtf8Bom's first three bytes
		// are the BOM.
		const std::string partName = SheetPart (parts, sheetName.empty ()
													   ? std::string () : Csv::ToUtf8Bom (sheetName).substr (3));
		if (partName.empty ())
			{
			why = L"the workbook has no sheet \"" + sheetName + L"\"";
			return false;
			}
		const auto sh = parts.find (partName);
		if (sh == parts.end ())
			{
			why = L"the workbook has no sheet (" + Csv::FromUtf8 (partName) + L")";
			return false;
			}
		const std::string &x = sh->second;

		// Every cell, by row number and column.
		size_t from = 0, rs = 0, re = 0, nextRow = 1;
		while (FindTag (x, "row", from, rs, re))
			{
			const std::string r = Attr (x, rs, re, "r");
			const size_t rowNo = r.empty () ? nextRow : static_cast<size_t> (std::strtoul (r.c_str (), nullptr, 10));
			nextRow = rowNo + 1;
			if (SelfClosing (x, re))
				{
				from = re;
				continue;
				}
			size_t rowEnd = 0;
			const std::string rowXml = Body (x, "row", re, rowEnd);
			from = rowEnd;

			size_t cf = 0, cs = 0, ce = 0, nextCol = 0;
			while (FindTag (rowXml, "c", cf, cs, ce))
				{
				size_t col = nextCol, rr = rowNo;
				const std::string ref = Attr (rowXml, cs, ce, "r");
				if (!ref.empty ())
					CellRef (ref, col, rr);
				nextCol = col + 1;
				if (SelfClosing (rowXml, ce))
					{
					cf = ce;
					continue;
					}
				size_t cellEnd = 0;
				const std::string cell = Body (rowXml, "c", ce, cellEnd);
				cf = cellEnd;

				const std::string t = Attr (rowXml, cs, ce, "t");
				std::wstring value;
				size_t vs = 0, ve = 0, vend = 0;
				if (t == "inlineStr")
					{
					if (FindTag (cell, "is", 0, vs, ve) && !SelfClosing (cell, ve))
						value = TextOf (Body (cell, "is", ve, vend));
					}
				else if (FindTag (cell, "v", 0, vs, ve) && !SelfClosing (cell, ve))
					{
					const std::string v = Body (cell, "v", ve, vend);
					if (t == "s")
						{
						const size_t idx = static_cast<size_t> (std::strtoul (v.c_str (), nullptr, 10));
						value = idx < shared.size () ? shared[idx] : std::wstring ();
						}
					else if (t.empty () || t == "n")
						{
						// A number. Excel writes up to 17 digits ("2.5000000000000001E-2");
						// the shortest text for the SAME double is what a person typed.
						value = Unescape (v);
						double d = 0;
						if (Csv::ParseDouble (value, d))
							value = Csv::FormatDouble (d);
						}
					else
						value = Unescape (v);	// "1"/"0", a formula's text, or "#N/A"
					}
				if (!value.empty ())
					grid[rowNo][col] = value;
				}
			}

		return true;
		}

	bool ReadGrid (const std::filesystem::path &file, const std::wstring &sheetName, Grid &grid, std::wstring &why)
		{
		std::ifstream in (file, std::ios::binary);
		if (!in)
			{
			why = L"could not be opened";
			return false;
			}
		const std::string bytes ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		return ReadGridBytes (bytes, sheetName, grid, why);
		}

	bool ReadNamedSheetBytes (const std::string &bytes, const std::wstring &sheetName,
							  std::vector<std::vector<std::wstring>> &rows,
							  std::vector<size_t> &sheetRow, std::wstring &why)
		{
		rows.clear ();
		sheetRow.clear ();
		Grid grid;
		if (!ReadGridBytes (bytes, sheetName, grid, why))
			return false;

		// From the column-name row down.
		size_t header = 0;
		for (const auto &r : grid)
			{
			for (const auto &c : r.second)
				if (c.second == L"op_idn")
					header = r.first;
			if (header != 0)
				break;
			}
		if (header == 0)
			{
			why = L"no column-name row (a row with \"op_idn\") - is this the dumped sheet?";
			return false;
			}

		size_t width = 0;
		for (const auto &c : grid[header])
			width = (std::max) (width, c.first + 1);
		for (const auto &r : grid)
			{
			if (r.first < header)
				continue;
			std::vector<std::wstring> row (width);
			bool any = false;
			for (const auto &c : r.second)
				if (c.first < width)
					{
					row[c.first] = c.second;
					any = any || !c.second.empty ();
					}
			if (!any)
				continue;
			rows.push_back (row);
			sheetRow.push_back (r.first);
			}
		return true;
		}

	bool ReadSheet (const std::filesystem::path &file,
					std::vector<std::vector<std::wstring>> &rows,
					std::vector<size_t> &sheetRow, std::wstring &why)
		{
		std::ifstream in (file, std::ios::binary);
		if (!in)
			{
			why = L"could not be opened";
			return false;
			}
		const std::string bytes ((std::istreambuf_iterator<char> (in)),
								 std::istreambuf_iterator<char> ());
		return ReadSheetBytes (bytes, rows, sheetRow, why);
		}

	bool ReadNamedSheet (const std::filesystem::path &file, const std::wstring &sheetName,
						 std::vector<std::vector<std::wstring>> &rows,
						 std::vector<size_t> &sheetRow, std::wstring &why)
		{
		std::ifstream in (file, std::ios::binary);
		if (!in)
			{
			why = L"could not be opened";
			return false;
			}
		const std::string bytes ((std::istreambuf_iterator<char> (in)),
								 std::istreambuf_iterator<char> ());
		return ReadNamedSheetBytes (bytes, sheetName, rows, sheetRow, why);
		}
	}
