#include "Xlsx.h"

#include <algorithm>
#include <cstdint>
#include <fstream>
#include <map>

namespace Xlsx
	{
	namespace
		{
		std::string Utf8 (const std::wstring &w)
			{
			std::string o;
			for (size_t i = 0; i < w.size (); ++i)
				{
				uint32_t c = static_cast<uint16_t> (w[i]);
				if (c >= 0xD800 && c < 0xDC00 && i + 1 < w.size ()
					&& static_cast<uint16_t> (w[i + 1]) >= 0xDC00
					&& static_cast<uint16_t> (w[i + 1]) < 0xE000)
					{
					c = 0x10000 + ((c - 0xD800) << 10)
						+ (static_cast<uint16_t> (w[++i]) - 0xDC00);
					}
				if (c < 0x80)
					o += static_cast<char> (c);
				else if (c < 0x800)
					{
					o += static_cast<char> (0xC0 | (c >> 6));
					o += static_cast<char> (0x80 | (c & 0x3F));
					}
				else if (c < 0x10000)
					{
					o += static_cast<char> (0xE0 | (c >> 12));
					o += static_cast<char> (0x80 | ((c >> 6) & 0x3F));
					o += static_cast<char> (0x80 | (c & 0x3F));
					}
				else
					{
					o += static_cast<char> (0xF0 | (c >> 18));
					o += static_cast<char> (0x80 | ((c >> 12) & 0x3F));
					o += static_cast<char> (0x80 | ((c >> 6) & 0x3F));
					o += static_cast<char> (0x80 | (c & 0x3F));
					}
				}
			return o;
			}

		/// XML text: escaped, and characters XML 1.0 forbids dropped.
		std::string Esc (const std::wstring &w)
			{
			std::wstring clean;
			for (wchar_t c : w)
				if (c == 9 || c == 10 || c == 13 || c >= 0x20)
					clean += c;
			std::string u = Utf8 (clean), o;
			for (char c : u)
				switch (c)
					{
					case '&': o += "&amp;"; break;
					case '<': o += "&lt;"; break;
					case '>': o += "&gt;"; break;
					case '"': o += "&quot;"; break;
					default: o += c;
					}
			return o;
			}

		}

	/// Text that is exactly a plain decimal number Excel would read back
	/// as the same text: -?digits[.digits], no leading zeros, no exponent.
	bool IsPlainNumber (const std::wstring &t)
		{
		{
			size_t i = 0;
			if (i < t.size () && t[i] == L'-')
				++i;
			const size_t d0 = i;
			while (i < t.size () && t[i] >= L'0' && t[i] <= L'9')
				++i;
			if (i == d0)
				return false;
			if (t[d0] == L'0' && i - d0 > 1)
				return false;
			if (i < t.size () && t[i] == L'.')
				{
				++i;
				const size_t f0 = i;
				while (i < t.size () && t[i] >= L'0' && t[i] <= L'9')
					++i;
				if (i == f0)
					return false;
				}
			return i == t.size () && t.size () < 24;
			}
		}

	namespace
		{

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

		void P16 (std::string &o, uint32_t v)
			{
			o += static_cast<char> (v & 0xFF);
			o += static_cast<char> ((v >> 8) & 0xFF);
			}
		void P32 (std::string &o, uint32_t v)
			{
			P16 (o, v & 0xFFFF);
			P16 (o, v >> 16);
			}

		// ---- Styles. Fills 0 and 1 are the two Excel requires.
		enum Fill { FNone = 0, FGray125 = 1, FGroup0 = 2 /* 8 of these */,
					FKind0 = 10 /* 5 */, FGrey = 15 };
		const char *const kGroupFills[8] = { "FFD9E1F2", "FFFCE4D6", "FFE2EFDA", "FFFFF2CC",
											 "FFDDEBF7", "FFEDEDED", "FFD0CECE", "FFE4DFEC" };
		const char *const kKindFills[5] = { "FFFFFFFF", "FFF2F8FC", "FFFCF6EC", "FFF3F9EE", "FFF6F1FA" };

		/// Fonts: 0 plain, 1 bold, 2 grey, 3 a page title, 4 a note, 5 a link,
		/// 6 a section heading.
		enum Font { NPlain = 0, NBold = 1, NGrey = 2, NTitle = 3, NNote = 4, NLink = 5, NSection = 6 };

		/// Borders: 0 none, 1 a box (a cell to type in on a report page), 2 a line
		/// under (a report table's heading).
		enum Border { BNone = 0, BBox = 1, BUnder = 2 };

		/// Registry of distinct cell formats -> cellXfs index.
		struct Styles
			{
			struct Key
				{
				int fill, font, fmt, border;
				bool right;
				bool operator< (const Key &o) const
					{
					if (fill != o.fill) return fill < o.fill;
					if (font != o.font) return font < o.font;
					if (fmt != o.fmt) return fmt < o.fmt;
					if (border != o.border) return border < o.border;
					return right < o.right;
					}
				};
			std::map<Key, int> ids;
			std::vector<Key> order;
			std::vector<std::string> custom;		//!< number formats 164 on
			Styles () { Get (0, 0); }
			/// fmt: a built-in number format - 0 General, 49 Text - or one of Fmt's.
			int Get (int fill, int font, int fmt = 0, int border = BNone, bool right = false)
				{
				const Key k = { fill, font, fmt, border, right };
				auto it = ids.find (k);
				if (it != ids.end ())
					return it->second;
				ids[k] = static_cast<int> (order.size ());
				order.push_back (k);
				return ids[k];
				}
			/// A number format's id: Excel's built-in where there is one (so a
			/// built-in follows the computer's own settings), else a custom one.
			int Fmt (const std::wstring &code)
				{
				if (code.empty ())
					return 0;
				if (code == L"currency")
					return 7;			// built in: the computer's own currency, 2 places
				static const std::pair<const wchar_t *, int> builtIn[] = {
					{ L"0", 1 }, { L"0.00", 2 }, { L"#,##0", 3 }, { L"#,##0.00", 4 }, { L"0%", 9 },
					{ L"0.00%", 10 }, { L"@", 49 } };
				for (const auto &b : builtIn)
					if (code == b.first)
						return b.second;
				const std::string c = Utf8 (code);
				for (size_t i = 0; i < custom.size (); ++i)
					if (custom[i] == c)
						return static_cast<int> (164 + i);
				custom.push_back (c);
				return static_cast<int> (164 + custom.size () - 1);
				}
			std::string Xml () const
				{
				std::string o =
					"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
					"<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
				if (!custom.empty ())
					{
					o += "<numFmts count=\"" + std::to_string (custom.size ()) + "\">";
					for (size_t i = 0; i < custom.size (); ++i)
						{
						std::string esc;
						for (char ch : custom[i])
							esc += ch == '"' ? std::string ("&quot;") : ch == '&' ? std::string ("&amp;")
								   : ch == '<' ? std::string ("&lt;") : ch == '>' ? std::string ("&gt;") : std::string (1, ch);
						o += "<numFmt numFmtId=\"" + std::to_string (164 + i) + "\" formatCode=\"" + esc + "\"/>";
						}
					o += "</numFmts>";
					}
				o +=
					"<fonts count=\"7\">"
					"<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
					"<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font>"
					"<font><sz val=\"11\"/><color rgb=\"FF7F7F7F\"/><name val=\"Calibri\"/></font>"
					"<font><b/><sz val=\"16\"/><color rgb=\"FF1F2937\"/><name val=\"Calibri\"/></font>"
					"<font><sz val=\"10\"/><color rgb=\"FF5B6573\"/><name val=\"Calibri\"/></font>"
					"<font><u/><sz val=\"11\"/><color rgb=\"FF0563C1\"/><name val=\"Calibri\"/></font>"
					"<font><b/><sz val=\"12\"/><color rgb=\"FF1F3864\"/><name val=\"Calibri\"/></font>"
					"</fonts><fills count=\"16\">"
					"<fill><patternFill patternType=\"none\"/></fill>"
					"<fill><patternFill patternType=\"gray125\"/></fill>";
				auto solid = [&o] (const char *rgb)
					{
					o += "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"";
					o += rgb;
					o += "\"/><bgColor indexed=\"64\"/></patternFill></fill>";
					};
				for (const char *c : kGroupFills)
					solid (c);
				for (const char *c : kKindFills)
					solid (c);
				solid ("FFBFBFBF");		// grey: not applicable / read-only
				o += "</fills><borders count=\"3\"><border><left/><right/><top/><bottom/>"
					 "<diagonal/></border>"
					 "<border><left style=\"thin\"><color rgb=\"FF8EA9DB\"/></left><right style=\"thin\"><color rgb=\"FF8EA9DB\"/>"
					 "</right><top style=\"thin\"><color rgb=\"FF8EA9DB\"/></top><bottom style=\"thin\"><color rgb=\"FF8EA9DB\"/>"
					 "</bottom><diagonal/></border>"
					 "<border><left/><right/><top/><bottom style=\"thin\"><color rgb=\"FF8EA9DB\"/></bottom><diagonal/></border>"
					 "</borders>"
					 "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" "
					 "borderId=\"0\"/></cellStyleXfs><cellXfs count=\"";
				o += std::to_string (order.size ()) + "\">";
				for (const Key &k : order)
					{
					o += "<xf numFmtId=\"" + std::to_string (k.fmt) + "\" fontId=\"" + std::to_string (k.font)
						 + "\" fillId=\"" + std::to_string (k.fill)
						 + "\" borderId=\"" + std::to_string (k.border) + "\" xfId=\"0\"";
					if (k.fmt != 0)
						o += " applyNumberFormat=\"1\"";
					if (k.fill != 0)
						o += " applyFill=\"1\"";
					if (k.font != 0)
						o += " applyFont=\"1\"";
					if (k.border != 0)
						o += " applyBorder=\"1\"";
					if (k.right)
						o += " applyAlignment=\"1\"><alignment horizontal=\"right\"/></xf>";
					else
						o += "/>";
					}
				o += "</cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" "
					 "xfId=\"0\" builtinId=\"0\"/></cellStyles>"
					 // 0: an edited cell; 1: the edit count of a row with edits;
					 // 2: a calculated cell that has moved with the edits
					 "<dxfs count=\"4\">"
					 "<dxf><font><b/><color rgb=\"FF6B3E00\"/></font><fill><patternFill patternType=\"solid\">"
					 "<bgColor rgb=\"FFFFE08A\"/></patternFill></fill></dxf>"
					 "<dxf><font><b/><color rgb=\"FF9A3412\"/></font><fill><patternFill patternType=\"solid\">"
					 "<bgColor rgb=\"FFFFD3A6\"/></patternFill></fill></dxf>"
					 "<dxf><font><b/><color rgb=\"FF1E3A8A\"/></font><fill><patternFill patternType=\"solid\">"
					 "<bgColor rgb=\"FFDBEAFE\"/></patternFill></fill></dxf>"
					 "<dxf><font><color rgb=\"FF9CA3AF\"/></font><fill><patternFill patternType=\"solid\">"
					 "<bgColor rgb=\"FFE5E7EB\"/></patternFill></fill></dxf>"
					 "</dxfs></styleSheet>";
				return o;
				}
			};
		}

	std::string ColName (size_t col)
		{
		std::string s;
		for (size_t n = col + 1; n > 0; n = (n - 1) / 26)
			s.insert (s.begin (), static_cast<char> ('A' + (n - 1) % 26));
		return s;
		}

	namespace
		{
		/// "A2:X9" -> "$A$2:$X$9".
		std::string AbsRef (const std::string &ref)
			{
			std::string o;
			bool letters = false;
			for (char c : ref)
				{
				const bool isLetter = c >= 'A' && c <= 'Z';
				const bool isDigit = c >= '0' && c <= '9';
				if (isLetter && !letters)
					o += '$';
				if (isDigit && letters)
					o += '$';
				letters = isLetter;
				o += c;
				}
			return o;
			}
		}

	std::string Zip (const std::vector<std::pair<std::string, std::string>> &parts)
		{
		std::string out, central;
		for (const auto &p : parts)
			{
			const uint32_t crc = Crc32 (p.second);
			const uint32_t size = static_cast<uint32_t> (p.second.size ());
			const uint32_t offset = static_cast<uint32_t> (out.size ());
			// DOS time 00:00:00, date 1980-01-01.
			auto common = [&] (std::string &o)
				{
				P16 (o, 20);			// version needed
				P16 (o, 0x0800);		// flags: UTF-8 names
				P16 (o, 0);				// stored
				P16 (o, 0);				// time
				P16 (o, 0x21);			// date
				P32 (o, crc);
				P32 (o, size);
				P32 (o, size);
				P16 (o, static_cast<uint32_t> (p.first.size ()));
				P16 (o, 0);				// extra
				};
			P32 (out, 0x04034B50);
			common (out);
			out += p.first;
			out += p.second;

			P32 (central, 0x02014B50);
			P16 (central, 20);			// version made by
			common (central);
			P16 (central, 0);			// comment length
			P16 (central, 0);			// disk
			P16 (central, 0);			// internal attrs
			P32 (central, 0);			// external attrs
			P32 (central, offset);
			central += p.first;
			}
		const uint32_t cdOffset = static_cast<uint32_t> (out.size ());
		out += central;
		P32 (out, 0x06054B50);
		P16 (out, 0);
		P16 (out, 0);
		P16 (out, static_cast<uint32_t> (parts.size ()));
		P16 (out, static_cast<uint32_t> (parts.size ()));
		P32 (out, static_cast<uint32_t> (central.size ()));
		P32 (out, cdOffset);
		P16 (out, 0);
		return out;
		}

	namespace
		{
		/// An attribute value: escaped, a line break kept as &#10; (a raw one
		/// would be read back as a space), and cut to Excel's limit.
		std::string Attr (std::wstring w, size_t limit)
			{
			if (w.size () > limit)
				w = w.substr (0, limit - 3) + L"...";
			std::string o;
			for (wchar_t c : w)
				{
				if (c == L'\r')
					continue;
				if (c == L'\n')
					o += "&#10;";
				else
					o += Esc (std::wstring (1, c));
				}
			return o;
			}
		}

	std::string Build (const Sheet &s)
		{
		const size_t nCols = s.rows.empty () ? 0 : s.rows[0].size ();
		Styles st;

		// Kind -> shading, in first-seen order.
		std::map<std::wstring, int> kindFill;
		for (size_t r = 1; r < s.rows.size (); ++r)
			{
			const std::wstring &k = s.rows[r][s.kindCol];
			if (!kindFill.count (k))
				kindFill[k] = FKind0 + static_cast<int> (kindFill.size () % 5);
			}

		auto flag = [] (const std::vector<char> &v, size_t i)
			{ return i < v.size () && v[i] != 0; };

		std::string sd;
		sd += "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
			  "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
		// Outline +/- buttons over a group's FIRST column (the one left showing).
		const bool macros = !s.vbaProject.empty ();
		sd += std::string ("<sheetPr") + (macros ? " codeName=\"Sheet1\"" : "")
			  + "><outlinePr summaryBelow=\"0\" summaryRight=\"0\"/></sheetPr>";

		// ---- Column outline: each outlined group keeps its first column at
		// level 0 (the handle) and puts the rest at level 1, so neighbouring
		// groups stay separate and each collapses on its own.
		std::vector<int> level (nCols, 0);
		std::vector<char> hidden (nCols, 0), collapsedHere (nCols, 0);
		bool anyOutline = false;
		for (size_t c = 0; c < nCols;)
			{
			const int g = c < s.group.size () ? s.group[c] : 0;
			size_t e = c;
			while (e + 1 < nCols && e + 1 < s.group.size () && s.group[e + 1] == g)
				++e;
			const bool outline = g >= 0 && static_cast<size_t> (g) < s.outlineGroup.size ()
								 && s.outlineGroup[static_cast<size_t> (g)] && e > c;
			if (outline)
				{
				const bool collapse = static_cast<size_t> (g) < s.collapseGroup.size ()
									  && s.collapseGroup[static_cast<size_t> (g)];
				for (size_t k = c + 1; k <= e; ++k)
					{
					level[k] = 1;
					hidden[k] = collapse;
					}
				collapsedHere[c] = collapse;
				anyOutline = true;
				}
			c = e + 1;
			}
		sd += "<sheetViews><sheetView workbookViewId=\"0\"><pane";
		if (s.frozenCols > 0)
			sd += " xSplit=\"" + std::to_string (s.frozenCols) + "\"";
		sd += " ySplit=\"2\" topLeftCell=\"" + ColName (s.frozenCols) + "3\" "
			  "activePane=\"" + std::string (s.frozenCols > 0 ? "bottomRight" : "bottomLeft")
			  + "\" state=\"frozen\"/>";
		if (s.frozenCols > 0)
			sd += "<selection pane=\"topRight\"/><selection pane=\"bottomLeft\"/>"
				  "<selection pane=\"bottomRight\" activeCell=\"" + ColName (s.frozenCols)
				  + "3\" sqref=\"" + ColName (s.frozenCols) + "3\"/>";
		else
			sd += "<selection pane=\"bottomLeft\"/>";
		sd += std::string ("</sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"15\"")
			  + (anyOutline ? " outlineLevelCol=\"1\"" : "") + "/><cols>";
		for (size_t c = 0; c < nCols; ++c)
			{
			size_t w = 8;
			for (const auto &row : s.rows)
				w = std::max (w, c < row.size () ? row[c].size () + 2 : 0);
			w = std::min<size_t> (w, c == 3 || flag (s.text, c) ? 40 : 30);
			sd += "<col min=\"" + std::to_string (c + 1) + "\" max=\""
				  + std::to_string (c + 1) + "\" width=\"" + std::to_string (w)
				  + "\" customWidth=\"1\""
				  + (level[c] ? " outlineLevel=\"1\"" : "")
				  + (hidden[c] ? " hidden=\"1\"" : "")
				  + (collapsedHere[c] ? " collapsed=\"1\"" : "") + "/>";
			}
		sd += "</cols><sheetData>";

		auto strCell = [&] (std::string &o, size_t r, size_t c, const std::wstring &t, int style)
			{
			o += "<c r=\"" + ColName (c) + std::to_string (r) + "\" s=\"" + std::to_string (style)
				 + "\"";
			if (t.empty ())
				{
				o += "/>";
				return;
				}
			o += " t=\"inlineStr\"><is><t xml:space=\"preserve\">" + Esc (t) + "</t></is></c>";
			};

		// Row 1: group names, merged across each group's columns.
		std::string merges;
		size_t nMerges = 0;
		sd += "<row r=\"1\">";
		for (size_t c = 0; c < nCols;)
			{
			const int g = c < s.group.size () ? s.group[c] : 0;
			size_t e = c;
			while (e + 1 < nCols && e + 1 < s.group.size () && s.group[e + 1] == g)
				++e;
			const int style = st.Get (FGroup0 + g % 8, 1);
			const std::wstring name = g < static_cast<int> (s.groupNames.size ())
										  ? s.groupNames[static_cast<size_t> (g)] : L"";
			for (size_t k = c; k <= e; ++k)
				strCell (sd, 1, k, k == c ? name : L"", style);
			if (e > c)
				{
				merges += "<mergeCell ref=\"" + ColName (c) + "1:" + ColName (e) + "1\"/>";
				++nMerges;
				}
			c = e + 1;
			}
		sd += "</row>";

		// Row 2: column names.
		sd += "<row r=\"2\">";
		for (size_t c = 0; c < nCols; ++c)
			{
			const int g = c < s.group.size () ? s.group[c] : 0;
			strCell (sd, 2, c, s.rows[0][c], st.Get (FGroup0 + g % 8, 1));
			}
		sd += "</row>";

		// ---- Edit tracking: where the rows are, and how a cell finds its dumped
		// value - by op_idn (column A), so a sorted sheet still compares right.
		const size_t lastRow = s.rows.size () + 1;
		auto dumpedCol = [&] (size_t c, bool relCol)
			{
			const std::string L = ColName (c);
			return "Dumped!" + std::string (relCol ? "" : "$") + L + "$3:" + (relCol ? "" : "$") + L
				   + "$" + std::to_string (lastRow);
			};
		auto match = [&] (size_t xr)
			{
			return "MATCH($A" + std::to_string (xr) + "," + dumpedCol (0, false) + ",0)";
			};
		const bool track = s.trackChanges && nCols > 0 && s.rows.size () > 1;
		const size_t changesCol = s.changesCol >= 0 ? static_cast<size_t> (s.changesCol) : nCols;

		// The column runs to compare: everything but the count itself.
		std::vector<std::pair<size_t, size_t>> runs;
		for (size_t c = 0; c < nCols;)
			{
			auto tracked = [&] (size_t k) { return k != changesCol && !flag (s.untracked, k); };
			if (!tracked (c))
				{
				++c;
				continue;
				}
			size_t e = c;
			while (e + 1 < nCols && tracked (e + 1))
				++e;
			runs.push_back ({ c, e });
			c = e + 1;
			}

		auto changesFormula = [&] (size_t xr)
			{
			std::string f;
			for (const auto &run : runs)
				{
				const std::string a = ColName (run.first), b = ColName (run.second);
				f += (f.empty () ? "" : "+");
				// (the Dumped row with this op_idn) x (its cells that differ). Plain
				// array maths: INDEX(...,0) here is cut to the cell's own column
				// by Excel's implicit intersection and always errors.
				const std::string n = std::to_string (lastRow), x = std::to_string (xr);
				f += "SUMPRODUCT((Dumped!$A$3:$A$" + n + "=$A" + x + ")*(Dumped!$" + a + "$3:$" + b
					 + "$" + n + "&lt;&gt;" + a + x + ":" + b + x + "))";
				}
			return "IFERROR(" + f + ",0)";
			};

		// Data rows.
		for (size_t r = 1; r < s.rows.size (); ++r)
			{
			const size_t xr = r + 2;
			const bool haveNa = r - 1 < s.notApplicable.size ();
			const int kind = kindFill[s.rows[r][s.kindCol]];
			sd += "<row r=\"" + std::to_string (xr) + "\">";
			for (size_t c = 0; c < nCols; ++c)
				{
				const std::wstring &t = c < s.rows[r].size () ? s.rows[r][c] : std::wstring ();
				const bool na = haveNa && flag (s.notApplicable[r - 1], c);
				const bool ro = flag (s.readOnly, c);
				const std::wstring fx = r - 1 < s.formula.size () && c < s.formula[r - 1].size ()
											? s.formula[r - 1][c] : std::wstring ();
				const bool derived = !fx.empty () && IsPlainNumber (t);
				const int style = na || ro || derived ? st.Get (FGrey, 2)
														: st.Get (kind, 0, flag (s.textFormat, c) ? 49 : 0);
				if (track && c == changesCol)
					sd += "<c r=\"" + ColName (c) + std::to_string (xr) + "\" s=\""
						  + std::to_string (st.Get (FGrey, 1)) + "\"><f>" + changesFormula (xr)
						  + "</f><v>0</v></c>";
				else if (derived)
					sd += "<c r=\"" + ColName (c) + std::to_string (xr) + "\" s=\""
						  + std::to_string (style) + "\"><f>" + Esc (fx) + "</f><v>" + Esc (t)
						  + "</v></c>";
				else if (!fx.empty ())
					// A formula that returns text.
					sd += "<c r=\"" + ColName (c) + std::to_string (xr) + "\" s=\""
						  + std::to_string (st.Get (FGrey, 2)) + "\" t=\"str\"><f>" + Esc (fx)
						  + "</f><v>" + Esc (t) + "</v></c>";
				else if (!t.empty () && !flag (s.text, c) && IsPlainNumber (t))
					sd += "<c r=\"" + ColName (c) + std::to_string (xr) + "\" s="
						  "\"" + std::to_string (style) + "\"><v>" + Esc (t) + "</v></c>";
				else
					strCell (sd, xr, c, t, style);
				}
			sd += "</row>";
			}
		sd += "</sheetData>";
		const std::string filterRef = "A2:" + ColName (nCols > 0 ? nCols - 1 : 0)
									  + std::to_string (s.rows.size () + 1);
		const bool filter = s.autoFilter && nCols > 0;
		if (filter)
			sd += "<autoFilter ref=\"" + filterRef + "\"/>";
		if (nMerges)
			sd += "<mergeCells count=\"" + std::to_string (nMerges) + "\">" + merges
				  + "</mergeCells>";
		// ---- Highlight every edited cell, and every row with edits.
		if (track)
			{
			int priority = 1;
			for (const auto &run : runs)
				{
				const std::string a = ColName (run.first), b = ColName (run.second);
				sd += "<conditionalFormatting sqref=\"" + a + "3:" + b + std::to_string (lastRow)
					  + "\"><cfRule type=\"expression\" dxfId=\"0\" priority=\""
					  + std::to_string (priority++) + "\"><formula>" + a + "3&lt;&gt;INDEX("
					  + dumpedCol (run.first, true) + "," + match (3) + ")</formula></cfRule>"
					  "</conditionalFormatting>";
				}
			// Calculated columns that follow the edits (the estimate, flips, MRR):
			// blue where they have moved from the dump - the effect of the edits.
			for (size_t c = 0; c < nCols; ++c)
				{
				if (c == changesCol || !flag (s.untracked, c))
					continue;
				const std::string a = ColName (c), cur = a + "3";
				const std::string was = "INDEX(" + dumpedCol (c, true) + "," + match (3) + ")";
				sd += "<conditionalFormatting sqref=\"" + a + "3:" + a + std::to_string (lastRow)
					  + "\"><cfRule type=\"expression\" dxfId=\"2\" priority=\"" + std::to_string (priority++)
					  + "\"><formula>IFERROR(IF(ISNUMBER(" + cur + "),ROUND(" + cur + ",6)&lt;&gt;ROUND(N(" + was
					  + "),6)," + cur + "&amp;\"\"&lt;&gt;" + was + "&amp;\"\"),FALSE)</formula></cfRule>"
					  "</conditionalFormatting>";
				}
			if (changesCol < nCols)
				sd += "<conditionalFormatting sqref=\"" + ColName (changesCol) + "3:" + ColName (changesCol)
					  + std::to_string (lastRow) + "\"><cfRule type=\"cellIs\" dxfId=\"1\" priority=\""
					  + std::to_string (priority++) + "\" operator=\"greaterThan\"><formula>0</formula>"
					  "</cfRule></conditionalFormatting>";
			}

		// ---- What each cell accepts, and its tooltip - for any page.
		auto validationsXml = [&] (const std::vector<Sheet::Validation> &rules)
		{
		std::string dv;
		size_t nDv = 0;
		for (const Sheet::Validation &v : rules)
			{
			if (v.cells.empty ())
				continue;
			std::string f1 = Esc (v.f1);
			if (v.type == "list")
				{
				std::wstring list;
				for (const std::wstring &c : v.choices)
					list += (list.empty () ? L"" : L",") + c;
				if (list.empty () || list.size () > 255)
					continue;					// too long for Excel - no rule, typing allowed
				f1 = "\"" + Esc (list) + "\"";
				}
			dv += "<dataValidation";
			if (!v.type.empty ())
				dv += " type=\"" + v.type + "\"";
			if (!v.op.empty ())
				dv += " operator=\"" + v.op + "\"";
			dv += " allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\"";
			if (!v.title.empty ())
				dv += " promptTitle=\"" + Attr (v.title, 32) + "\" errorTitle=\"" + Attr (v.title, 32) + "\"";
			if (!v.prompt.empty ())
				dv += " prompt=\"" + Attr (v.prompt, 255) + "\"";
			if (!v.error.empty ())
				dv += " error=\"" + Attr (v.error, 225) + "\"";
			dv += " sqref=\"" + v.cells + "\">";
			if (!v.type.empty ())
				{
				dv += "<formula1>" + f1 + "</formula1>";
				if (!v.f2.empty ())
					dv += "<formula2>" + Esc (v.f2) + "</formula2>";
				}
			dv += "</dataValidation>";
			++nDv;
			}
		return nDv ? "<dataValidations count=\"" + std::to_string (nDv) + "\">" + dv + "</dataValidations>"
				   : std::string ();
		};
		sd += validationsXml (s.validations);

		// Tool numbers link to their row on the Tools page.
		const bool toolsPage = !s.tools.empty ();
		if (toolsPage && !s.toolLinks.empty ())
			{
			sd += "<hyperlinks>";
			for (const auto &l : s.toolLinks)
				if (l.second < s.tools.size ())
					sd += "<hyperlink ref=\"" + l.first + "\" location=\"'Tools'!A"
						  + std::to_string (l.second + 2) + "\" display=\""
						  + Esc (s.tools[l.second].number) + "\"/>";
			sd += "</hyperlinks>";
			}
		sd += "</worksheet>";

		const std::string decl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";

		// ---- "Dumped": the values exactly as written, cell for cell, for the
		// highlighting to compare against. Hidden; never read by a load.
		std::string dumped;
		if (track)
			{
			dumped = decl + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>";
			for (size_t r = 0; r < s.rows.size (); ++r)
				{
				const size_t xr = r + 2;
				dumped += "<row r=\"" + std::to_string (xr) + "\">";
				for (size_t c = 0; c < nCols; ++c)
					{
					std::wstring t = c < s.rows[r].size () ? s.rows[r][c] : std::wstring ();
					if (r > 0 && c == changesCol)
						t = L"0";
					if (t.empty ())
						continue;
					const std::string ref = ColName (c) + std::to_string (xr);
					if (r > 0 && (c == changesCol || (!flag (s.text, c) && IsPlainNumber (t))))
						dumped += "<c r=\"" + ref + "\"><v>" + Esc (t) + "</v></c>";
					else
						dumped += "<c r=\"" + ref + "\" t=\"inlineStr\"><is><t xml:space=\"preserve\">"
								  + Esc (t) + "</t></is></c>";
					}
				dumped += "</row>";
				}
			dumped += "</sheetData></worksheet>";
			}

		// ---- Free cells (the Tools page's extras, the Summary page): text, a number
		// or a formula, styled by what the cell is for.
		auto cell = [&] (size_t r, size_t c, const std::wstring &t, int style)
			{
			std::string o = "<c r=\"" + ColName (c) + std::to_string (r) + "\" s=\"" + std::to_string (style) + "\"";
			if (t.empty ())
				return o + "/>";
			return o + " t=\"inlineStr\"><is><t xml:space=\"preserve\">" + Esc (t) + "</t></is></c>";
			};
		auto freeCell = [&] (size_t r, size_t c, const Sheet::FreeCell &f)
			{
			const int fmt = f.textFormat ? 49 : st.Fmt (f.numFmt);
			int style = 0;
			switch (f.look)
				{
				case Sheet::FreeCell::Title: style = st.Get (FNone, NTitle, fmt); break;
				case Sheet::FreeCell::Section: style = st.Get (FNone, NSection, fmt); break;
				case Sheet::FreeCell::Note: style = st.Get (FNone, NNote, fmt, BNone, f.right); break;
				case Sheet::FreeCell::Plain:
					style = st.Get (FNone, f.head ? NBold : NPlain, fmt, f.head ? BUnder : BNone, f.right);
					break;
				case Sheet::FreeCell::Link: style = st.Get (FNone, NLink, fmt, BNone, f.right); break;
				case Sheet::FreeCell::Input: style = st.Get (FNone, NBold, fmt, BBox, f.right); break;
				default:
					style = f.head ? st.Get (FGroup0, NBold, fmt) : f.editable ? st.Get (FNone, NPlain, fmt)
																			  : st.Get (FGrey, NGrey, fmt);
				}
			const std::string ref = ColName (c) + std::to_string (r);
			const std::string at = "<c r=\"" + ref + "\" s=\"" + std::to_string (style) + "\"";
			if (!f.formula.empty ())
				{
				const std::string fx = f.array ? "<f t=\"array\" ref=\"" + ref + "\">" + Esc (f.formula) + "</f>"
											   : "<f>" + Esc (f.formula) + "</f>";
				return IsPlainNumber (f.text)
						   ? at + ">" + fx + "<v>" + Esc (f.text) + "</v></c>"
						   : at + " t=\"str\">" + fx + "<v>" + Esc (f.text) + "</v></c>";
				}
			if (!f.head && !f.text.empty () && IsPlainNumber (f.text))
				return at + "><v>" + Esc (f.text) + "</v></c>";
			return cell (r, c, f.text, style);
			};

		// ---- "Summary": the page the workbook opens on - what to look at first.
		const bool summaryPage = !s.summary.empty ();
		std::string summaryXml;
		if (summaryPage)
			{
			// Printed one page wide, as long as it needs.
			summaryXml = decl + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
						 "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><sheetViews><sheetView showGridLines=\"0\" tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>"
						 "<sheetFormatPr defaultRowHeight=\"15\"/>";
			if (!s.summaryWidths.empty ())
				{
				summaryXml += "<cols>";
				for (size_t c = 0; c < s.summaryWidths.size (); ++c)
					{
					const double w = s.summaryWidths[c];
					summaryXml += "<col min=\"" + std::to_string (c + 1) + "\" max=\"" + std::to_string (c + 1)
								  + "\" width=\"" + std::to_string (w > 0 ? w : 8) + "\" customWidth=\"1\""
								  + (w > 0 ? "" : " hidden=\"1\"") + "/>";
					}
				summaryXml += "</cols>";
				}
			summaryXml += "<sheetData>";
			for (size_t k = 0; k < s.summary.size (); ++k)
				{
				const size_t r = k + 1;
				summaryXml += "<row r=\"" + std::to_string (r) + "\">";
				for (size_t c = 0; c < s.summary[k].size (); ++c)
					{
					const Sheet::FreeCell &f = s.summary[k][c];
					// An empty plain cell is no cell at all: text runs on over it.
					if (f.text.empty () && f.formula.empty () && f.look != Sheet::FreeCell::Auto
						&& f.look != Sheet::FreeCell::Input && !f.head)
						continue;
					summaryXml += freeCell (r, c, f);
					}
				summaryXml += "</row>";
				}
			summaryXml += "</sheetData>" + validationsXml (s.summaryValidations) + "<pageMargins left=\"0.5\" right=\"0.5\" top=\"0.6\" bottom=\"0.6\" "
						  "header=\"0.3\" footer=\"0.3\"/><pageSetup orientation=\"landscape\" fitToWidth=\"1\" fitToHeight=\"0\"/>"
						  "</worksheet>";
			}

		// ---- "Tools": number, name, which operations use it, and its picture.
		std::string toolsXml, drawingXml, drawingRels;
		std::vector<std::pair<std::string, std::string>> media;
		if (toolsPage)
			{
			const int boxW = 220, boxH = 145;		// the picture's room, pixels
			const std::string ns = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
								   "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
			const size_t nExtra = s.toolExtraHeads.size ();
			const size_t picCol = 3 + nExtra;
			toolsXml = decl + "<worksheet " + ns + ">"
					   "<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" "
					   "activePane=\"bottomLeft\" state=\"frozen\"/></sheetView></sheetViews>"
					   "<sheetFormatPr defaultRowHeight=\"15\"/><cols>"
					   "<col min=\"1\" max=\"1\" width=\"8\" customWidth=\"1\"/>"
					   "<col min=\"2\" max=\"2\" width=\"40\" customWidth=\"1\"/>"
					   "<col min=\"3\" max=\"3\" width=\"30\" customWidth=\"1\"/>";
			if (nExtra > 0)
				toolsXml += "<col min=\"4\" max=\"" + std::to_string (3 + nExtra)
							+ "\" width=\"16\" customWidth=\"1\"/>";
			toolsXml += "<col min=\"" + std::to_string (picCol + 1) + "\" max=\"" + std::to_string (picCol + 1)
						+ "\" width=\"33\" customWidth=\"1\"/></cols><sheetData>";
			const int head = st.Get (FGroup0, 1);
			toolsXml += "<row r=\"1\">" + cell (1, 0, L"Tool", head) + cell (1, 1, L"Name", head)
						+ cell (1, 2, L"Used by", head);
			for (size_t e = 0; e < nExtra; ++e)
				toolsXml += cell (1, 3 + e, s.toolExtraHeads[e], head);
			toolsXml += cell (1, picCol, L"Picture", head) + "</row>";

			std::string anchors;
			for (size_t k = 0; k < s.tools.size (); ++k)
				{
				const Sheet::ToolRow &t = s.tools[k];
				const size_t r = k + 2;
				const bool pic = !t.png.empty () && t.width > 0 && t.height > 0;
				toolsXml += "<row r=\"" + std::to_string (r) + "\""
							+ std::string (pic ? " ht=\"115\" customHeight=\"1\"" : "") + ">"
							+ cell (r, 0, t.number, 0) + cell (r, 1, t.name, 0) + cell (r, 2, t.usedBy, 0);
				for (size_t e = 0; e < t.extra.size () && e < nExtra; ++e)
					toolsXml += freeCell (r, 3 + e, t.extra[e]);
				toolsXml += "</row>";
				if (!pic)
					continue;
				const size_t n = media.size () + 1;
				media.push_back ({ "xl/media/image" + std::to_string (n) + ".png", t.png });
				const double scale = (std::min) (static_cast<double> (boxW) / t.width,
												  static_cast<double> (boxH) / t.height);
				const long long cx = static_cast<long long> (t.width * scale * 9525.0);
				const long long cy = static_cast<long long> (t.height * scale * 9525.0);
				anchors += "<xdr:oneCellAnchor><xdr:from><xdr:col>" + std::to_string (picCol) + "</xdr:col><xdr:colOff>38100</xdr:colOff>"
						   "<xdr:row>" + std::to_string (r - 1) + "</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:from>"
						   "<xdr:ext cx=\"" + std::to_string (cx) + "\" cy=\"" + std::to_string (cy) + "\"/>"
						   "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + std::to_string (n + 1) + "\" name=\"Tool "
						   + Esc (t.number) + "\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr>"
						   "</xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId" + std::to_string (n) + "\"/>"
						   "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\">"
						   "<a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
				drawingRels += "<Relationship Id=\"rId" + std::to_string (n) + "\" Type=\"http://schemas.openxmlformats.org/"
							   "officeDocument/2006/relationships/image\" Target=\"../media/image" + std::to_string (n)
							   + ".png\"/>";
				}
			for (size_t k = 0; k < s.toolsAfter.size (); ++k)
				{
				const size_t r = s.tools.size () + 3 + k;
				toolsXml += "<row r=\"" + std::to_string (r) + "\">";
				for (size_t c = 0; c < s.toolsAfter[k].size (); ++c)
					toolsXml += freeCell (r, c, s.toolsAfter[k][c]);
				toolsXml += "</row>";
				}
			toolsXml += "</sheetData>";
			// Greyed: an input that does not apply on that row (dxf 3).
			for (size_t g = 0; g < s.toolsGreyed.size (); ++g)
				toolsXml += "<conditionalFormatting sqref=\"" + s.toolsGreyed[g].cells + "\"><cfRule type=\"expression\" "
							"dxfId=\"3\" priority=\"" + std::to_string (g + 1) + "\"><formula>"
							+ Esc (s.toolsGreyed[g].formula) + "</formula></cfRule></conditionalFormatting>";
			toolsXml += validationsXml (s.toolsValidations);
			if (!media.empty ())
				{
				toolsXml += "<drawing r:id=\"rId1\"/>";
				drawingXml = decl + "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
							 "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
							 "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
							 + anchors + "</xdr:wsDr>";
				drawingRels = decl + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
							  + drawingRels + "</Relationships>";
				}
			toolsXml += "</worksheet>";
			}

		std::vector<std::pair<std::string, std::string>> parts;
		parts.push_back ({ "[Content_Types].xml", decl +
			"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
			"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
			"<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
			"<Default Extension=\"png\" ContentType=\"image/png\"/>"
			+ std::string (macros ? "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>" : "") +
			"<Override PartName=\"/xl/workbook.xml\" ContentType=\""
			+ std::string (macros ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
								  : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") + "\"/>"
			"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
			+ std::string (track ? "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" : "")
			+ std::string (toolsPage ? "<Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" : "")
			+ std::string (summaryPage ? "<Override PartName=\"/xl/worksheets/sheet4.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" : "")
			+ std::string (!drawingXml.empty () ? "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>" : "") +
			"<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
			"</Types>" });
		parts.push_back ({ "_rels/.rels", decl +
			"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
			"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
			+ std::string (macros && !s.ribbonXml.empty ()
							   ? "<Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2007/relationships/ui/extensibility\" Target=\"customUI/customUI14.xml\"/>"
							   : "")
			+ "</Relationships>" });
		parts.push_back ({ "xl/workbook.xml", decl +
			"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
			"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
			+ std::string (macros ? "<workbookPr codeName=\"ThisWorkbook\"/>" : "") +
			// The Summary first: the page it opens on. Code names stay as they were -
			// the macros find their sheets by name, and the main sheet by Sheet1.
			"<bookViews><workbookView activeTab=\"0\"/></bookViews><sheets>"
			+ std::string (summaryPage ? "<sheet name=\"Summary\" sheetId=\"4\" r:id=\"rId6\"/>" : "")
			+ "<sheet name=\"Lathe params\" sheetId=\"1\" r:id=\"rId1\"/>"
			+ std::string (toolsPage ? "<sheet name=\"Tools\" sheetId=\"3\" r:id=\"rId4\"/>" : "")
			+ std::string (track ? "<sheet name=\"Dumped\" sheetId=\"2\" state=\"hidden\" r:id=\"rId3\"/>" : "")
			+ "</sheets>"
			// localSheetId is the main sheet's POSITION, which the Summary moves.
			+ (filter || !s.helpFolder.empty () ? std::string ("<definedNames>") : std::string ())
			+ (filter ? "<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\""
						+ std::string (summaryPage ? "1" : "0") + "\" "
						"hidden=\"1\">'Lathe params'!" + AbsRef (filterRef) + "</definedName>"
					  : std::string ())
			// Where the user manual is: the macros' help buttons open its pages.
			+ (!s.helpFolder.empty ()
				   ? "<definedName name=\"PT_Help\" hidden=\"1\">&quot;" + Esc (s.helpFolder) + "&quot;</definedName>"
				   : std::string ())
			+ (filter || !s.helpFolder.empty () ? std::string ("</definedNames>") : std::string ())
			// Recalculate on open: the counts and highlights are formulas.
			+ "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>"
			+ "</workbook>" });
		parts.push_back ({ "xl/_rels/workbook.xml.rels", decl +
			"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
			"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
			"<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
			+ std::string (track ? "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>" : "")
			+ std::string (toolsPage ? "<Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>" : "")
			+ std::string (summaryPage ? "<Relationship Id=\"rId6\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet4.xml\"/>" : "")
			+ std::string (macros ? "<Relationship Id=\"rId5\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>" : "")
			+ "</Relationships>" });
		parts.push_back ({ "xl/styles.xml", st.Xml () });
		parts.push_back ({ "xl/worksheets/sheet1.xml", sd });
		if (track)
			parts.push_back ({ "xl/worksheets/sheet2.xml", dumped });
		if (macros)
			{
			parts.push_back ({ "xl/vbaProject.bin", s.vbaProject });
			if (!s.ribbonXml.empty ())
				parts.push_back ({ "customUI/customUI14.xml", s.ribbonXml });
			}
		if (summaryPage)
			parts.push_back ({ "xl/worksheets/sheet4.xml", summaryXml });
		if (toolsPage)
			{
			parts.push_back ({ "xl/worksheets/sheet3.xml", toolsXml });
			if (!drawingXml.empty ())
				{
				parts.push_back ({ "xl/worksheets/_rels/sheet3.xml.rels", decl +
					"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
					"<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
					"Target=\"../drawings/drawing1.xml\"/></Relationships>" });
				parts.push_back ({ "xl/drawings/drawing1.xml", drawingXml });
				parts.push_back ({ "xl/drawings/_rels/drawing1.xml.rels", drawingRels });
				for (const auto &m : media)
					parts.push_back (m);
				}
			}
		return Zip (parts);
		}

	bool Write (const std::filesystem::path &file, const Sheet &s)
		{
		const std::string bytes = Build (s);
		std::ofstream f (file, std::ios::binary | std::ios::trunc);
		if (!f)
			return false;
		f.write (bytes.data (), static_cast<std::streamsize> (bytes.size ()));
		f.close ();
		return !f.fail ();
		}
	}
