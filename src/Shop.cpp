#include "Shop.h"
#include "Csv.h"
#include "Inspect.h"

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <map>
#include <set>
#include <tuple>

namespace
	{
	const double kPi = 3.14159265358979323846;

	std::wstring Lower (const std::wstring &s)
		{
		std::wstring o;
		for (wchar_t c : s)
			o += static_cast<wchar_t> (std::towlower (c));
		return o;
		}

	/// A figure to show, rounded to `places`; "" for none.
	std::wstring Num (double v, int places)
		{
		if (std::isnan (v))
			return std::wstring ();
		const double k = std::pow (10.0, places);
		return Csv::Tidy (std::round (v * k) / k);
		}

	/// Seconds as h:mm:ss; "" for none.
	std::wstring Hms (double seconds)
		{
		if (!(seconds >= 0))
			return std::wstring ();
		const long long t = static_cast<long long> (seconds + 0.5);
		wchar_t buf[32];
		swprintf_s (buf, L"%lld:%02lld:%02lld", t / 3600, (t / 60) % 60, t % 60);
		return buf;
		}

	/// The op numbers in a Tools row's "Used by" ("op 2, 5, 12").
	std::vector<std::wstring> IdsIn (const std::wstring &usedBy)
		{
		std::vector<std::wstring> ids;
		std::wstring n;
		for (size_t i = 0; i <= usedBy.size (); ++i)
			if (i < usedBy.size () && std::iswdigit (usedBy[i]))
				n += usedBy[i];
			else if (!n.empty ())
				{
				ids.push_back (n);
				n.clear ();
				}
		return ids;
		}

	/// Text joined with ", ".
	std::wstring Join (const std::vector<std::wstring> &items, const wchar_t *sep = L", ")
		{
		std::wstring o;
		for (const std::wstring &s : items)
			o += (o.empty () ? L"" : sep) + s;
		return o;
		}
	}

namespace Shop
	{
	double EdgeSeconds (double cut, double flips, double longest, double partsPerEdge)
		{
		if (!(cut > 0))
			return kNone;
		if (flips >= 1)
			{
			// The average edge, unless one edge is known to have run longer - the
			// worn edge is the one that matters.
			const double e = cut / flips;
			return longest > e ? longest : e;
			}
		if (partsPerEdge > 0)
			return cut * partsPerEdge;
		return kNone;			// it outlasts a part, by how much nobody said
		}

	Spread SpreadOf (std::vector<double> values)
		{
		values.erase (std::remove_if (values.begin (), values.end (), [] (double v) { return std::isnan (v); }),
					  values.end ());
		Spread s;
		s.n = values.size ();
		if (values.empty ())
			return s;
		std::sort (values.begin (), values.end ());
		s.min = values.front ();
		s.max = values.back ();
		const size_t m = values.size () / 2;
		s.median = values.size () % 2 == 1 ? values[m] : (values[m - 1] + values[m]) / 2.0;
		return s;
		}

	void Read (const Xlsx::Sheet &s, Part &p)
		{
		p.ops.clear ();
		if (s.rows.empty ())
			return;
		const std::vector<std::wstring> &names = s.rows[0];
		auto colOf = [&names] (const std::wstring &n) -> int
			{
			for (size_t c = 0; c < names.size (); ++c)
				if (names[c] == n)
					return static_cast<int> (c);
			return -1;
			};
		auto text = [&colOf] (const std::vector<std::wstring> &row, const std::wstring &name)
			{
			const int c = colOf (name);
			return c >= 0 && static_cast<size_t> (c) < row.size () ? Csv::Trim (row[static_cast<size_t> (c)]) : std::wstring ();
			};
		auto num = [&text] (const std::vector<std::wstring> &row, const std::wstring &name)
			{
			double v = 0;
			return Csv::ParseDouble (text (row, name), v) ? v : kNone;
			};

		// ---- Each op's tool, as the Tools page lists it (a row per tool, its ops
		// in "Used by") - so two tools sharing a number stay apart - and its insert:
		// the cell headed "Insert", else the first one (the dump's own order).
		size_t insAt = 0;
		for (size_t e = 0; e < s.toolExtraHeads.size (); ++e)
			if (s.toolExtraHeads[e] == L"Insert")
				{
				insAt = e;
				break;
				}
		std::map<std::wstring, size_t> toolOfOp;			// op_idn -> Tools row
		for (size_t k = 0; k < s.tools.size (); ++k)
			for (const std::wstring &id : IdsIn (s.tools[k].usedBy))
				toolOfOp[id] = k;
		auto insertOf = [&s, insAt] (size_t k)
			{
			return insAt < s.tools[k].extra.size () ? Csv::Trim (s.tools[k].extra[insAt].text) : std::wstring ();
			};

		// ---- The inserts table under the tools: per insert its edges, flips, cost
		// and parts per edge, found by their headings. It runs to the first empty
		// row; rows nobody named are the spare ones for typing in.
		struct Info
			{
			double edges = kNone, flips = kNone, cost = kNone, partsPerEdge = 0;
			};
		std::map<std::wstring, Info> table;					// by insert, lower case
		for (size_t h = 0; h < s.toolsAfter.size (); ++h)
			{
			const auto &head = s.toolsAfter[h];
			if (head.empty () || !head[0].head || head[0].text != L"Inserts")
				continue;
			auto at = [&head] (const wchar_t *heading) -> int
				{
				for (size_t c = 0; c < head.size (); ++c)
					if (head[c].text == heading)
						return static_cast<int> (c);
				return -1;
				};
			const int cName = at (L"Insert"), cEdges = at (L"Edges per insert"), cFlips = at (L"Flips / part"),
					  cCost = at (L"Cost per insert"), cPpe = at (L"Parts per edge");
			for (size_t r = h + 1; r < s.toolsAfter.size () && !s.toolsAfter[r].empty (); ++r)
				{
				const auto &row = s.toolsAfter[r];
				auto cell = [&row] (int c) { return c >= 0 && static_cast<size_t> (c) < row.size () ? Csv::Trim (row[static_cast<size_t> (c)].text) : std::wstring (); };
				const std::wstring name = cell (cName);
				if (name.empty ())
					continue;
				Info in;
				double v = 0;
				if (Csv::ParseDouble (cell (cEdges), v))
					in.edges = v;
				if (Csv::ParseDouble (cell (cFlips), v))
					in.flips = v;
				if (Csv::ParseDouble (cell (cCost), v))
					in.cost = v;
				if (Csv::ParseDouble (cell (cPpe), v) && v > 0)
					in.partsPerEdge = v;
				table[Lower (name)] = in;
				}
			break;
			}
		// Insert cost per part as the Summary page works it out: flips / edges x cost,
		// over the inserts with a cost (and edges) typed.
		p.inserts = table.size ();
		p.costed = 0;
		p.insertCost = kNone;
		for (const auto &kv : table)
			if (kv.second.edges > 0 && !std::isnan (kv.second.cost))
				{
				++p.costed;
				if (!std::isnan (kv.second.flips))
					p.insertCost = (std::isnan (p.insertCost) ? 0.0 : p.insertCost) + kv.second.flips / kv.second.edges * kv.second.cost;
				}

		// ---- The rows.
		struct Sum
			{
			double cut = 0, flips = 0, longest = 0;
			};
		std::map<size_t, Sum> sums;							// per Tools row
		std::vector<int> rowTool;							// per op: its Tools row, or -1
		double cycle = 0, cut = 0, flips = 0;
		bool anyCycle = false, anyCut = false, anyFlips = false;
		p.regen = 0;
		p.units.clear ();
		for (size_t r = 1; r < s.rows.size (); ++r)
			{
			const std::vector<std::wstring> &row = s.rows[r];
			Op o;
			o.part = p.name;
			o.op = text (row, L"op_idn");
			if (o.op.empty ())
				continue;							// not an op's row
			o.type = text (row, L"type");
			o.comment = text (row, L"comment");
			o.units = text (row, L"units");
			o.needsRegen = text (row, L"needs_regen");
			if (o.needsRegen == L"yes")
				++p.regen;
			if (p.units.empty ())
				p.units = o.units;
			const auto tk = toolOfOp.find (o.op);
			const int k = tk == toolOfOp.end () ? -1 : static_cast<int> (tk->second);
			rowTool.push_back (k);
			if (k >= 0)
				{
				o.tool = s.tools[static_cast<size_t> (k)].number;
				o.insert = insertOf (static_cast<size_t> (k));
				}
			const bool mm = o.units == L"mm";

			// The speed and feed that cut. A PrimeTurning op keeps its own, per pass:
			// the rough pass's, else the finish pass's.
			o.speed = text (row, L"speed");
			o.speedMode = text (row, L"speed_mode");
			o.feed = text (row, L"feed");
			o.feedMode = text (row, L"feed_mode");
			if (o.type == L"PRIME")
				for (const wchar_t *pass : { L"rough", L"fin" })
					{
					const std::wstring pre = std::wstring (L"pt_") + pass + L"_";
					const std::wstring on = text (row, std::wstring (pass) == L"rough" ? L"do_rough" : L"do_finish");
					const std::wstring speed = text (row, pre + L"speed");
					if (on == L"0" || speed.empty ())
						continue;
					o.speed = speed;
					o.speedMode = text (row, pre + L"css") == L"1" ? L"CSS" : L"RPM";
					o.feed = text (row, pre + L"feed_axial");
					o.feedMode = text (row, pre + L"axial_type");
					break;
					}

			// The depth of cut: the column the dump's own MRR works from.
			const wchar_t *ap = nullptr;
			if (o.type == L"ROUGH" || o.type == L"PRIME")
				ap = L"step";
			else if (o.type == L"FINISH" && num (row, L"n_cuts") > 1)
				ap = L"step";						// a finish's step is a depth only between passes
			else if (o.type == L"DYNAMIC")
				ap = L"stepover";
			else if (!text (row, L"rough_step").empty ())
				ap = L"rough_step";					// face, groove, plunge rough
			if (ap != nullptr)
				o.depth = text (row, ap);

			// Surface speed at the op's mean cutting diameter (cut_dia - a drill's is its
			// own), CSS held to max_ss there; feed per rev from a feed per minute at that
			// RPM. As the dump's MRR works them out.
			const double k1 = mm ? 1000.0 : 12.0;
			double speed = 0, feed = 0, depth = 0;
			const bool hasSpeed = Csv::ParseDouble (o.speed, speed) && speed > 0;
			const bool hasFeed = Csv::ParseDouble (o.feed, feed) && feed > 0;
			const double dia = num (row, L"cut_dia"), cap = num (row, L"max_ss");
			double rpm = kNone, vc = kNone;
			if (hasSpeed && o.speedMode == L"CSS")
				{
				if (dia > 0)
					{
					rpm = speed * k1 / (kPi * dia);
					if (cap > 0)
						rpm = (std::min) (rpm, cap);
					vc = rpm * kPi * dia / k1;
					}
				else
					vc = speed;
				}
			else if (hasSpeed)
				{
				rpm = speed;
				if (dia > 0)
					vc = speed * kPi * dia / k1;
				}
			double perRev = kNone;
			if (hasFeed && o.feedMode == L"per rev")
				perRev = feed;
			else if (hasFeed && o.feedMode == L"per min" && rpm > 0)
				perRev = feed / rpm;
			// To inch units: m/min to ft/min, mm to in.
			const double toIn = mm ? 1.0 / 25.4 : 1.0;
			o.sfm = vc * (mm ? 1000.0 / 304.8 : 1.0);
			o.feedIpr = perRev * toIn;
			if (Csv::ParseDouble (o.depth, depth) && depth > 0)
				o.depthIn = depth * toIn;

			o.estSeconds = num (row, L"est_seconds");
			o.cutSeconds = num (row, L"cut_seconds_est");
			o.flips = num (row, L"flips_part");
			o.mrr = num (row, L"mrr");
			o.airPct = num (row, L"air_pct");
			if (!std::isnan (o.mrr))
				o.mrrUnits = mm ? L"cm3/min" : L"in3/min";

			// The part's totals, as its Summary page adds them up.
			if (!std::isnan (o.estSeconds))
				{
				cycle += o.estSeconds;
				anyCycle = true;
				}
			if (!std::isnan (o.cutSeconds))
				{
				cut += o.cutSeconds;
				anyCut = true;
				}
			if (!std::isnan (o.flips))
				{
				flips += o.flips;
				anyFlips = true;
				}

			// Its tool's: each op and its copies by transforms (as the Tools page sums).
			if (k >= 0)
				{
				const double copies = num (row, L"xf_copies");
				const double w = 1 + (copies > 0 ? copies : 0);
				Sum &sum = sums[static_cast<size_t> (k)];
				if (!std::isnan (o.cutSeconds))
					sum.cut += o.cutSeconds * w;
				if (!std::isnan (o.flips))
					sum.flips += o.flips * w;
				double longest = 0;
				if (Inspect::ParseMinSec (text (row, L"flip_longest"), longest))
					sum.longest = (std::max) (sum.longest, longest);
				}
			p.ops.push_back (o);
			}
		p.cycleSeconds = anyCycle ? cycle : kNone;
		p.cutSeconds = anyCut ? cut : kNone;
		p.flips = anyFlips ? flips : kNone;

		// Each op gets its tool's edge time.
		for (size_t i = 0; i < p.ops.size () && i < rowTool.size (); ++i)
			{
			if (rowTool[i] < 0)
				continue;
			const Sum &sum = sums[static_cast<size_t> (rowTool[i])];
			const auto in = table.find (Lower (p.ops[i].insert));
			const double ppe = in == table.end () ? 0.0 : in->second.partsPerEdge;
			p.ops[i].edgeSeconds = EdgeSeconds (sum.cut, sum.flips, sum.longest, ppe);
			}
		}

	std::vector<InsertRow> ByInsert (const std::vector<Part> &parts)
		{
		using Key = std::tuple<std::wstring, std::wstring, std::wstring>;
		struct Gather
			{
			InsertRow row;
			std::vector<double> sfm, feed, depth, edge;
			std::set<std::pair<std::wstring, std::wstring>> edgeOf;		// (part, tool): its edge counted once
			std::vector<std::pair<std::wstring, std::vector<std::wstring>>> tools;	// per part, its tools
			};
		std::map<Key, Gather> groups;
		for (const Part &p : parts)
			{
			if (!p.ok)
				continue;
			for (const Op &o : p.ops)
				{
				if (o.insert.empty ())
					continue;
				Gather &g = groups[Key (Lower (o.insert), Lower (o.type), Lower (p.material))];
				if (g.row.nOps == 0)
					{
					g.row.insert = o.insert;
					g.row.opKind = o.type;
					g.row.material = p.material;
					}
				++g.row.nOps;
				if (std::find (g.row.parts.begin (), g.row.parts.end (), p.name) == g.row.parts.end ())
					g.row.parts.push_back (p.name);
				const std::wstring tool = L"T" + o.tool;
				auto pt = std::find_if (g.tools.begin (), g.tools.end (), [&p] (const auto &x) { return x.first == p.name; });
				if (pt == g.tools.end ())
					{
					g.tools.push_back ({ p.name, {} });
					pt = g.tools.end () - 1;
					}
				if (std::find (pt->second.begin (), pt->second.end (), tool) == pt->second.end ())
					pt->second.push_back (tool);
				g.sfm.push_back (o.sfm);
				g.feed.push_back (o.feedIpr);
				g.depth.push_back (o.depthIn);
				if (!std::isnan (o.edgeSeconds) && g.edgeOf.insert ({ p.name, o.tool }).second)
					g.edge.push_back (o.edgeSeconds);
				if (!std::isnan (o.flips))
					g.row.flips += o.flips;
				}
			}
		std::vector<InsertRow> out;
		for (auto &kv : groups)
			{
			Gather &g = kv.second;
			g.row.sfm = SpreadOf (g.sfm);
			g.row.feed = SpreadOf (g.feed);
			g.row.depth = SpreadOf (g.depth);
			g.row.edge = SpreadOf (g.edge);
			std::vector<std::wstring> each;
			for (const auto &t : g.tools)
				each.push_back (t.first + L": " + Join (t.second));
			g.row.tools = Join (each, L"; ");
			out.push_back (g.row);
			}
		return out;
		}

	std::vector<std::filesystem::path> FindParts (const std::filesystem::path &folder, bool subfolders,
												  const std::wstring &ext, size_t limit, bool &more)
		{
		more = false;
		std::vector<std::filesystem::path> out;
		std::error_code ec;
		if (folder.empty () || !std::filesystem::is_directory (folder, ec))
			return out;
		const std::wstring want = Lower (ext);
		// False = stop looking.
		auto take = [&] (const std::filesystem::directory_entry &e)
			{
			std::error_code fe;
			if (!e.is_regular_file (fe) || Lower (e.path ().extension ().wstring ()) != want)
				return true;
			if (limit > 0 && out.size () >= limit)
				{
				more = true;
				return false;
				}
			out.push_back (e.path ());
			return true;
			};
		if (subfolders)
			{
			std::filesystem::recursive_directory_iterator it (folder, std::filesystem::directory_options::skip_permission_denied, ec);
			const std::filesystem::recursive_directory_iterator end;
			while (!ec && it != end)
				{
				if (!take (*it))
					break;
				it.increment (ec);
				}
			}
		else
			for (const auto &e : std::filesystem::directory_iterator (folder, ec))
				if (!take (e))
					break;
		std::sort (out.begin (), out.end (), [] (const std::filesystem::path &a, const std::filesystem::path &b)
			{ return _wcsicmp (a.c_str (), b.c_str ()) < 0; });
		return out;
		}

	std::wstring FileName (std::time_t when)
		{
		std::tm tmv = {};
		localtime_s (&tmv, &when);
		wchar_t date[16] = L"";
		std::wcsftime (date, 16, L"%Y%m%d", &tmv);
		return L"Shop_compare_" + std::wstring (date) + L".xlsx";
		}

	std::vector<Xlsx::Table> Tables (const std::vector<Part> &parts, const std::wstring &title)
		{
		std::vector<Xlsx::Table> book;

		// ---- Parts.
		{
		Xlsx::Table t;
		t.name = L"Parts";
		t.rows.push_back ({ L"part", L"file", L"status", L"why", L"material", L"units", L"ops", L"ops_needing_regen",
							L"cycle_time", L"cycle_seconds", L"cut_seconds", L"flips_part", L"inserts", L"inserts_costed",
							L"insert_cost_part", L"kinds_not_read", L"workbook" });
		t.text = { 1, 1, 1, 1, 1, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 1, 1 };
		for (const Part &p : parts)
			t.rows.push_back ({ p.name, p.file, p.ok ? L"dumped" : L"not dumped", p.why, p.material, p.units,
								p.ok ? std::to_wstring (p.ops.size ()) : L"", p.ok ? std::to_wstring (p.regen) : L"",
								Hms (p.cycleSeconds), Num (p.cycleSeconds, 1), Num (p.cutSeconds, 1), Num (p.flips, 2),
								p.ok ? std::to_wstring (p.inserts) : L"",
								p.ok ? std::to_wstring (p.costed) + L" of " + std::to_wstring (p.inserts) : L"",
								Num (p.insertCost, 2), p.skipped, p.workbook });
		book.push_back (t);
		}

		// ---- Inserts: the names another tool reads (insert, op_kind, material,
		// sfm_median, feed_median, depth_median, edge_time_median, n_ops, parts)
		// must not change.
		{
		Xlsx::Table t;
		t.name = L"Inserts";
		t.rows.push_back ({ L"insert", L"op_kind", L"material", L"n_ops", L"parts", L"part_list", L"tools",
							L"sfm_min", L"sfm_median", L"sfm_max", L"feed_min", L"feed_median", L"feed_max",
							L"depth_min", L"depth_median", L"depth_max", L"edge_time_min", L"edge_time_median",
							L"edge_time_max", L"edge_time_n", L"flips" });
		t.text = { 1, 1, 1, 0, 0, 1, 1 };
		for (const InsertRow &r : ByInsert (parts))
			t.rows.push_back ({ r.insert, r.opKind, r.material, std::to_wstring (r.nOps), std::to_wstring (r.parts.size ()),
								Join (r.parts), r.tools,
								Num (r.sfm.min, 1), Num (r.sfm.median, 1), Num (r.sfm.max, 1),
								Num (r.feed.min, 5), Num (r.feed.median, 5), Num (r.feed.max, 5),
								Num (r.depth.min, 4), Num (r.depth.median, 4), Num (r.depth.max, 4),
								Num (r.edge.min, 0), Num (r.edge.median, 0), Num (r.edge.max, 0), std::to_wstring (r.edge.n),
								Num (r.flips, 2) });
		book.push_back (t);
		}

		// ---- Ops.
		{
		Xlsx::Table t;
		t.name = L"Ops";
		t.rows.push_back ({ L"part", L"op", L"type", L"comment", L"tool", L"insert", L"material", L"units", L"speed",
							L"speed_mode", L"feed", L"feed_mode", L"depth", L"sfm", L"feed_ipr", L"depth_in", L"est_time",
							L"est_seconds", L"cut_seconds", L"flips_part", L"edge_seconds", L"mrr", L"mrr_units", L"air_pct",
							L"needs_regen" });
		t.text = { 1, 0, 1, 1, 0, 1, 1, 1, 0, 1, 0, 1, 0, 0, 0, 0, 1 };
		for (const Part &p : parts)
			for (const Op &o : p.ops)
				t.rows.push_back ({ o.part, o.op, o.type, o.comment, o.tool, o.insert, p.material, o.units, o.speed,
									o.speedMode, o.feed, o.feedMode, o.depth, Num (o.sfm, 1), Num (o.feedIpr, 5),
									Num (o.depthIn, 4), Hms (o.estSeconds), Num (o.estSeconds, 1), Num (o.cutSeconds, 1),
									Num (o.flips, 2), Num (o.edgeSeconds, 0), Num (o.mrr, 3), o.mrrUnits, Num (o.airPct, 1),
									o.needsRegen });
		book.push_back (t);
		}

		// ---- About: every column, in words.
		{
		Xlsx::Table t;
		t.name = L"About";
		t.widths = { 10, 20, 110 };
		t.rows.push_back ({ L"sheet", L"column", L"what it holds" });
		t.rows.push_back ({ L"", L"", title });
		t.rows.push_back ({ L"", L"", L"Each part was opened (not regenerated, never saved), dumped as \"Lathe params - dump to Excel\" "
									  L"dumps a whole part, and closed. Its own workbook is named under Parts > workbook." });
		t.rows.push_back ({ L"", L"", L"Speeds, feeds and depths on Inserts (and sfm, feed_ipr, depth_in on Ops) are in INCH units: "
									  L"a metric part's m/min, mm/rev and mm are converted. Times are in seconds." });
		const wchar_t *const about[][3] = {
			{ L"Parts", L"part", L"The part file's name." },
			{ L"Parts", L"file", L"Where it is, from the folder dumped." },
			{ L"Parts", L"status / why", L"dumped, or not dumped and why (it would not open, no ops this tool reads ...)." },
			{ L"Parts", L"material", L"The stock material of its machine group, as Mastercam names it (blank: none set)." },
			{ L"Parts", L"ops_needing_regen", L"Ops that needed regenerating when dumped - their times and flips are from the old toolpath." },
			{ L"Parts", L"cycle_time / cycle_seconds", L"The estimated cycle time: its ops' est_seconds added up, as the part's own Summary page." },
			{ L"Parts", L"cut_seconds", L"Time spent feeding (cutting), all ops." },
			{ L"Parts", L"flips_part", L"Insert flips per part, all ops." },
			{ L"Parts", L"inserts / inserts_costed", L"Inserts the part uses, and how many have a cost typed on its Tools page (kept in its .ptconfig)." },
			{ L"Parts", L"insert_cost_part", L"Insert cost per part (flips / edges x cost) over the inserts with a cost. Blank: none costed." },
			{ L"Parts", L"kinds_not_read", L"Kinds of op the tool does not read yet - left out of every figure." },
			{ L"Parts", L"workbook", L"The part's own dump, written beside it (or in the folder chosen)." },
			{ L"Inserts", L"insert", L"The insert, as the part's Tools page names it (a name typed over the guess there wins)." },
			{ L"Inserts", L"op_kind", L"The kind of op: ROUGH, FINISH, GROOVE, FACE, DRILL, PRIME ..." },
			{ L"Inserts", L"material", L"The stock material (blank: unknown). One row per insert, kind of op and material." },
			{ L"Inserts", L"n_ops / parts / part_list", L"How many ops, how many parts, and which parts." },
			{ L"Inserts", L"tools", L"Which tools carry it, per part." },
			{ L"Inserts", L"sfm_min / _median / _max", L"Surface speed, feet per minute, at each op's mean cutting diameter (CSS held to max RPM there)." },
			{ L"Inserts", L"feed_min / _median / _max", L"Feed per rev, inches. A per-minute feed is divided by the op's RPM." },
			{ L"Inserts", L"depth_min / _median / _max", L"Depth of cut, inches: step (rough, PrimeTurning, a finish of several passes), stepover (dynamic), rough step (face, groove, plunge rough)." },
			{ L"Inserts", L"edge_time_min / _median / _max", L"Seconds one edge cuts before it is turned: the tool's cut time per part over its flips per part (or the longest between flips, if longer); with parts per edge typed, cut time x parts per edge. Each tool counted once." },
			{ L"Inserts", L"edge_time_n", L"How many tools gave an edge time (an edge that outlasts the part, with no parts per edge typed, gives none)." },
			{ L"Inserts", L"flips", L"Insert flips per part, added over these ops (one of each part)." },
			{ L"Ops", L"part / op / type / comment / tool / insert", L"The op, as on the part's own sheet." },
			{ L"Ops", L"speed / speed_mode", L"As dumped, in the part's units: CSS = surface speed (SFM, or m/min metric), RPM = spindle speed. PrimeTurning: its rough pass's (else finish's)." },
			{ L"Ops", L"feed / feed_mode", L"As dumped, in the part's units: per rev or per min." },
			{ L"Ops", L"depth", L"Depth of cut as dumped, in the part's units (blank: not a roughing kind)." },
			{ L"Ops", L"sfm / feed_ipr / depth_in", L"The same in inch units - what Inserts works from. Blank: cannot be said (an RPM op with no cutting diameter, a per-minute feed with no RPM)." },
			{ L"Ops", L"est_time / est_seconds", L"Estimated time of the op (Mastercam's, as dumped)." },
			{ L"Ops", L"cut_seconds", L"Time spent feeding." },
			{ L"Ops", L"flips_part", L"Insert flips per part in this op." },
			{ L"Ops", L"edge_seconds", L"Its tool's edge time, in seconds (see Inserts > edge_time)." },
			{ L"Ops", L"mrr / mrr_units", L"Removal rate while cutting (in3/min, or cm3/min metric) - as the part's sheet." },
			{ L"Ops", L"air_pct", L"Share of the cutting time spent in air (from the stock simulation, where it ran)." },
			{ L"Ops", L"needs_regen", L"yes: the op needed regenerating when dumped." } };
		for (const auto &a : about)
			t.rows.push_back ({ a[0], a[1], a[2] });
		t.text = { 1, 1, 1 };
		book.push_back (t);
		}
		return book;
		}
	}
