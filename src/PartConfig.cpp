#include "PartConfig.h"
#include "Csv.h"
#include "Xlsx.h"

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <fstream>
#include <iterator>
#include <system_error>

namespace
	{
	std::wstring Trim (const std::wstring &s)
		{
		const size_t a = s.find_first_not_of (L" \t\r\n");
		if (a == std::wstring::npos)
			return std::wstring ();
		return s.substr (a, s.find_last_not_of (L" \t\r\n") - a + 1);
		}

	std::wstring Lower (std::wstring s)
		{
		for (wchar_t &c : s)
			c = static_cast<wchar_t> (std::towlower (c));
		return s;
		}

	/// A cell of the grid, "" when empty.
	std::wstring At (const Xlsx::Grid &g, size_t row, size_t col)
		{
		const auto r = g.find (row);
		if (r == g.end ())
			return std::wstring ();
		const auto c = r->second.find (col);
		return c == r->second.end () ? std::wstring () : Trim (c->second);
		}

	void Keep (std::wstring &to, const std::wstring &v)
		{
		if (!v.empty ())
			to = v;
		}

	/// A tool's figure as it stands in the workbook against what the dump filled
	/// in: the same (a number to a hair - Excel saves 9.525 as 9.5250000000000004),
	/// or blank where the dump left it blank ("-").
	bool SameFigure (const std::wstring &cell, const std::wstring &detected)
		{
		if (cell.empty ())
			return detected.empty () || detected == L"-";
		double a = 0, b = 0;
		if (Csv::ParseDouble (cell, a) && Csv::ParseDouble (detected, b))
			return std::fabs (a - b) <= 1e-9 * (std::max) (1.0, std::fabs (b));
		return cell == detected;
		}
	}

namespace PartConfig
	{
	std::filesystem::path PathFor (const std::filesystem::path &part)
		{
		return part.parent_path () / (part.stem ().wstring () + L".ptconfig");
		}

	Config Load (const std::filesystem::path &file)
		{
		Config c;
		std::ifstream in (file, std::ios::binary);
		if (!in)
			return c;
		const std::string bytes ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		const std::wstring text = Csv::FromUtf8 (bytes.compare (0, 3, "\xEF\xBB\xBF") == 0 ? bytes.substr (3) : bytes);
		std::wstring section;
		size_t at = 0;
		while (at <= text.size ())
			{
			size_t end = text.find (L'\n', at);
			if (end == std::wstring::npos)
				end = text.size ();
			const std::wstring line = Trim (text.substr (at, end - at));
			at = end + 1;
			if (line.empty () || line[0] == L'#' || line[0] == L';')
				continue;
			if (line.front () == L'[' && line.back () == L']')
				{
				section = Trim (line.substr (1, line.size () - 2));
				continue;
				}
			const size_t eq = line.find (L'=');
			if (eq == std::wstring::npos)
				continue;
			const std::wstring key = Lower (Trim (line.substr (0, eq))), value = Trim (line.substr (eq + 1));
			if (section == L"batch" && key == L"qty")
				c.batchQty = value;
			else if (section.rfind (L"insert ", 0) == 0)
				{
				Insert &i = c.inserts[Trim (section.substr (7))];
				if (key == L"edges") i.edges = value;
				else if (key == L"cost") i.cost = value;
				else if (key == L"parts_per_edge") i.partsPerEdge = value;
				else if (key == L"edges_detected") i.edgesDetected = value;
				}
			else if (section.rfind (L"tool ", 0) == 0)
				{
				Tool &t = c.tools[Trim (section.substr (5))];
				if (key == L"edge_life") t.edgeLife = value;
				else if (key == L"insert") t.insert = value;
				else if (key == L"detected") t.detected = value;
				else if (key == L"life_detected") t.lifeDetected = value;
				else if (key == L"ic") t.ic = value;
				else if (key == L"ic_detected") t.icDetected = value;
				else if (key == L"entering_angle") t.angle = value;
				else if (key == L"angle_detected") t.angleDetected = value;
				}
			}
		return c;
		}

	bool Save (const std::filesystem::path &file, const Config &c)
		{
		std::wstring t = L"# Parameter Table Tool - what was typed into this part's dumps that Mastercam does\r\n"
						 L"# not hold. A new dump starts from it. Safe to edit or delete.\r\n";
		if (!c.batchQty.empty ())
			t += L"\r\n[batch]\r\nqty = " + c.batchQty + L"\r\n";
		for (const auto &kv : c.inserts)
			{
			const Insert &i = kv.second;
			if (i.edges.empty () && i.cost.empty () && i.partsPerEdge.empty () && i.edgesDetected.empty ())
				continue;
			t += L"\r\n[insert " + kv.first + L"]\r\n";
			if (!i.edges.empty ()) t += L"edges = " + i.edges + L"\r\n";
			if (!i.cost.empty ()) t += L"cost = " + i.cost + L"\r\n";
			if (!i.partsPerEdge.empty ()) t += L"parts_per_edge = " + i.partsPerEdge + L"\r\n";
			if (!i.edgesDetected.empty ()) t += L"edges_detected = " + i.edgesDetected + L"\r\n";
			}
		for (const auto &kv : c.tools)
			{
			const Tool &tl = kv.second;
			if (tl.edgeLife.empty () && tl.insert.empty () && tl.detected.empty () && tl.lifeDetected.empty ()
				&& tl.ic.empty () && tl.icDetected.empty () && tl.angle.empty () && tl.angleDetected.empty ())
				continue;
			t += L"\r\n[tool " + kv.first + L"]\r\n";
			if (!tl.edgeLife.empty ()) t += L"edge_life = " + tl.edgeLife + L"\r\n";
			if (!tl.insert.empty ()) t += L"insert = " + tl.insert + L"\r\n";
			if (!tl.detected.empty ()) t += L"detected = " + tl.detected + L"\r\n";
			if (!tl.lifeDetected.empty ()) t += L"life_detected = " + tl.lifeDetected + L"\r\n";
			if (!tl.ic.empty ()) t += L"ic = " + tl.ic + L"\r\n";
			if (!tl.icDetected.empty ()) t += L"ic_detected = " + tl.icDetected + L"\r\n";
			if (!tl.angle.empty ()) t += L"entering_angle = " + tl.angle + L"\r\n";
			if (!tl.angleDetected.empty ()) t += L"angle_detected = " + tl.angleDetected + L"\r\n";
			}
		std::ofstream out (file, std::ios::binary | std::ios::trunc);
		if (!out)
			return false;
		const std::string bytes = Csv::ToUtf8Bom (t);
		out.write (bytes.data (), static_cast<std::streamsize> (bytes.size ()));
		return static_cast<bool> (out);
		}

	bool Harvest (const std::filesystem::path &workbook, Config &c, std::wstring &why)
		{
		Xlsx::Grid tools, summary;
		if (!Xlsx::ReadGrid (workbook, L"Tools", tools, why))
			return false;
		Config n = c;
		// The tool rows: a header row 1 (Tool | Name | Used by | Insert | ...), the
		// tools under it to the first row with no tool number. The columns are
		// found by their headings, so a later layout still reads.
		size_t insCol = 3, lifeCol = 7;
		size_t icCol = 0, angleCol = 0;			// 0: an older workbook without them
		for (const auto &h : tools[1])
			{
			const std::wstring head = Lower (Trim (h.second));
			if (head == L"insert") insCol = h.first;
			else if (head.rfind (L"edge life", 0) == 0) lifeCol = h.first;
			else if (head.rfind (L"insert size", 0) == 0) icCol = h.first;
			else if (head.rfind (L"entering angle", 0) == 0) angleCol = h.first;
			}
		for (size_t r = 2; !At (tools, r, 0).empty (); ++r)
			{
			const std::wstring number = At (tools, r, 0);
			Tool &t = n.tools[number];
			// Kept only when it differs from what the dump filled in.
			if (!t.lifeDetected.empty ())
				{
				const std::wstring life = At (tools, r, lifeCol);
				t.edgeLife = life != t.lifeDetected ? life : std::wstring ();
				}
			const std::wstring typed = At (tools, r, insCol);
			// A name other than what the dump put there is a correction. Without a
			// record of what it put there (a workbook from before this file), the
			// name cannot be told from the dump's own label: left alone.
			if (!t.detected.empty () && !typed.empty ())
				t.insert = typed != t.detected ? typed : std::wstring ();
			// Insert size and entering angle: like the edge life - kept only when they
			// differ from what that dump filled in (blank: back to the dump's own).
			if (icCol > 0 && !t.icDetected.empty ())
				{
				const std::wstring ic = At (tools, r, icCol);
				t.ic = SameFigure (ic, t.icDetected) ? std::wstring () : ic;
				}
			if (angleCol > 0 && !t.angleDetected.empty ())
				{
				const std::wstring angle = At (tools, r, angleCol);
				t.angle = SameFigure (angle, t.angleDetected) ? std::wstring () : angle;
				}
			}
		// The inserts table: the row whose first cell says "Inserts", then a row
		// per insert (name in B) to the first blank row; headings found by name.
		for (const auto &row : tools)
			{
			if (At (tools, row.first, 0) != L"Inserts")
				continue;
			size_t edgesCol = 3, costCol = 6, partsCol = 8;
			for (const auto &h : row.second)
				{
				const std::wstring head = Lower (Trim (h.second));
				if (head.rfind (L"edges", 0) == 0) edgesCol = h.first;
				else if (head == L"cost per insert") costCol = h.first;
				else if (head.rfind (L"parts per edge", 0) == 0) partsCol = h.first;
				}
			for (size_t r = row.first + 1; tools.count (r); ++r)
				{
				const std::wstring name = At (tools, r, 1);
				if (name.empty ())
					continue;
				Insert &i = n.inserts[name];
				const std::wstring edges = At (tools, r, edgesCol);
				if (!i.edgesDetected.empty () || i.edges.empty ())
					i.edges = edges != i.edgesDetected ? edges : std::wstring ();
				// Every dump writes the kept cost and parts per edge back into these
				// cells, so a blank one here was cleared on purpose.
				i.cost = At (tools, r, costCol);
				i.partsPerEdge = At (tools, r, partsCol);
				}
			break;
			}
		// The batch quantity, beside its label on the Summary.
		std::wstring sumWhy;
		if (Xlsx::ReadGrid (workbook, L"Summary", summary, sumWhy))
			for (const auto &row : summary)
				for (const auto &cell : row.second)
					if (Lower (Trim (cell.second)).rfind (L"batch quantity", 0) == 0)
						Keep (n.batchQty, At (summary, row.first, cell.first + 1));
		c = n;
		return true;
		}

	std::wstring TypedOr (const std::wstring &own, const std::wstring &typed, std::wstring &detected)
		{
		const std::wstring v = !detected.empty () && !typed.empty () ? typed : own;
		detected = own.empty () ? std::wstring (L"-") : own;
		return v;
		}

	std::filesystem::path NewestDump (const std::filesystem::path &part, const std::filesystem::path &folder)
		{
		const std::wstring stem = Lower (part.stem ().wstring ());
		std::filesystem::path best;
		std::filesystem::file_time_type bestTime {};
		for (const std::filesystem::path &dir : { part.parent_path (), folder })
			{
			std::error_code ec;
			if (dir.empty () || !std::filesystem::is_directory (dir, ec))
				continue;
			for (const auto &e : std::filesystem::directory_iterator (dir, ec))
				{
				const std::wstring name = Lower (e.path ().filename ().wstring ());
				const std::wstring ext = Lower (e.path ().extension ().wstring ());
				if ((ext != L".xlsx" && ext != L".xlsm") || name.rfind (stem, 0) != 0
					|| name.find (L"_params") == std::wstring::npos || name.rfind (L"~$", 0) == 0)
					continue;
				const auto t = std::filesystem::last_write_time (e.path (), ec);
				if (!ec && (best.empty () || t > bestTime))
					{
					best = e.path ();
					bestTime = t;
					}
				}
			}
		return best;
		}
	}
