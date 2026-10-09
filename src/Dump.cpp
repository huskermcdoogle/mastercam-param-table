#include "stdafx.h"
#include "MastercamSdk.h"
#include "Dump.h"
#include "LatheFields.h"
#include "Util.h"
#include "Csv.h"
#include "Xlsx.h"
#include "Coolant.h"
#include "ColumnHelp.h"
#include "Paths.h"
#include "StockSim.h"
#include "Estimate.h"
#include "DumpDialog.h"
#include "Settings.h"
#include "Inspect.h"

#include <cwctype>
#include "FileRules.h"
#include "ToolPictures.h"
#include "resource.h"
#include "SetupSheet_CH.h"

#include <algorithm>
#include <cmath>
#include <ctime>
#include <cwchar>
#include <fstream>
#include <map>
#include <memory>
#include <set>

namespace
	{
	/// THE OPERATION'S OWN COMMENT - what somebody typed on it, and what the Operation
	/// Manager shows. NOT op_make_description, which builds Mastercam's own generic
	/// text about the operation and is not what anyone recognises their job by.
	std::wstring Comment (operation *pOp)
		{
		return std::wstring (pOp->comment, wcsnlen (pOp->comment, COMMENT_SIZE));
		}

	/// THE TOOL RADIUS the percent columns are a percent OF. The SDK says "percent
	/// of tool radius" without saying which of the tool's numbers that is (corner
	/// radius or half the diameter), so pick the one that reproduces the stored
	/// amount / percent pair. `how` says how it was chosen, for the log.
	enum class Radius { Corner, HalfDia, Guess };

	double ToolRadius (const operation *pOp, double amount, double percent, Radius &how)
		{
		const double corner = pOp->tl.crad;
		const double half = pOp->tl.dia / 2.0;
		auto fits = [&] (double r)
			{ return r > 0 && percent > 0 && std::fabs (amount / r * 100.0 - percent) < 1e-6 * percent; };
		if (fits (corner))
			{ how = Radius::Corner; return corner; }
		if (fits (half))
			{ how = Radius::HalfDia; return half; }
		how = Radius::Guess;
		return corner > 0 ? corner : half;
		}

	/// The toolpath group's name, as the Operation Manager shows it.
	std::wstring GroupName (long id)
		{
		TpGrpList &groups = TpMainGrpMgr.GetMainGrpList ();
		const INT_PTR i = groups.IndexByID (id);
		if (i < 0)
			return std::wstring ();
		const op_group *g = groups.GetAt (i);
		return g == nullptr ? std::wstring ()
							: std::wstring (g->name, wcsnlen (g->name, MAX_GROUP_NAME + 1));
		}

	/// The operation's canned text codes, raw - where X-style coolant is expected
	/// to live. Shown undecoded until real operations with known coolant settings
	/// say what the codes mean.
	std::wstring CannedText (const op_canned_text &ct)
		{
		std::wstring s;
		for (short code : ct.cantxt)
			if (code != 0)
				s += (s.empty () ? L"" : L" ") + std::to_wstring (code);
		if (!ct.on && !s.empty ())
			s = L"(off) " + s;
		return s;
		}

	/// Everything the writers need to know about one column of the sheet.
	struct Column
		{
		std::wstring name;
		int group = 0;
		bool readOnly = false;
		bool text = false;
		bool isDouble = false;
		bool info = false;			//!< written for the reader, not read from a table
		};

	/// The groups, in the order they sit on the sheet.
	const std::vector<std::wstring> kGroupNames = {
		L"Identity", L"Feeds and speeds", L"Depth of cut", L"Stock to leave", L"Coolant",
		L"Toolpath", L"Home position", L"Reference points", L"Planes", L"Filter",
		L"Where it sits (read-only)", L"Tool inspection", L"Toolpath stats (read-only)",
		L"Insert flips (read-only)", L"Material removal (read-only)", L"Working numbers (read-only)" };
	enum { GIdentity, GFeeds, GDepth, GStock, GCoolant, GToolpath,
		   GHome, GRef, GPlanes, GFilter, GWhere, GInspect, GStats, GFlips, GMrr, GCalc };

	/// Groups with a +/- of their own, and the ones that start folded away -
	/// rarely edited, still one click from view.
	const int kOutlined[] = { GFeeds, GDepth, GStock, GCoolant, GStats, GFlips, GMrr, GCalc, GToolpath, GHome,
							  GRef, GPlanes, GFilter, GWhere, GInspect };
	const int kCollapsed[] = { GRef, GFilter, GWhere, GInspect, GCalc };

	/// Seconds as h:mm:ss.
	std::wstring Hms (double seconds)
		{
		if (!(seconds >= 0))
			return std::wstring ();
		const long long t = static_cast<long long> (seconds + 0.5);
		wchar_t buf[32];
		swprintf_s (buf, L"%lld:%02lld:%02lld", t / 3600, (t / 60) % 60, t % 60);
		return buf;
		}

	/// THE TOOLPATH'S OWN NUMBERS, from Mastercam: estimated cycle time and how
	/// far the tool travels. Read-only, and stale while the operation needs
	/// regenerating - needs_regen says so.
	struct Stats
		{
		std::wstring time, timeRaw, xMin, xMax, zMin, zMax, cut, rapid;
		double seconds = 0;
		Paths::Totals path;

		/// Tool inspection: the op's settings and what its NCI did with them.
		bool inspOk = false;
		Inspect::Settings insp;
		Inspect::Result inspRes;

		double toolDia = 0;			//!< for a drill's MRR

		/// What the 2D stock simulation says the op removed (StockSim.h).
		bool simOk = false, simHasAir = false, simFromBoundary = false, simBoundaryBad = false, simAgrees = true;
		double simRemoved = 0, simAirPct = 0, simOwn = -1;
		std::wstring simCheck;
		};

	Stats StatsOf (operation *pOp)
		{
		Stats st;
		const unsigned long diag = Settings::Diag ();
		const double t = (diag & 1) ? 0.0 : CalcCycleTime (pOp, false);
		st.seconds = t;
		if (!(diag & 2))
			st.path = Paths::Walk (*pOp);
		if (st.path.ok)
			{
			st.cut = Csv::Tidy (std::round (st.path.cutLength * 10000.0) / 10000.0);
			st.rapid = Csv::Tidy (std::round (st.path.rapidLength * 10000.0) / 10000.0);
			}
		if (t > 0)
			{
			st.time = Hms (t);
			st.timeRaw = Csv::Tidy (t);
			}
		// Travel from our own walk of the NCI. NOT GetOperationExtents: it runs
		// the setup-sheet tool report, which on a part with two different tools
		// sharing a number asks about it - once per operation dumped.
		if (st.path.ok && st.path.moves > 0)
			{
			auto tidy = [] (double v) { return Csv::Tidy (std::round (v * 10000.0) / 10000.0); };
			st.xMin = tidy (st.path.min[0]);
			st.xMax = tidy (st.path.max[0]);
			st.zMin = tidy (st.path.min[2]);
			st.zMax = tidy (st.path.max[2]);
			}
		return st;
		}

	/// The Plan column behind a sheet column (types and limits are the same in
	/// every table that has it), or nullptr.
	const Plan::Col *PlanCol (const std::wstring &name)
		{
		for (const Lathe::Table &t : Lathe::AllTables ())
			{
			const int at = Lathe::IndexOf (t, name);
			if (at >= 0)
				return &t.schema.cols[static_cast<size_t> (at)];
			}
		return nullptr;
		}

	/// What a cell of this column accepts, in words - the tooltip's last line
	/// and the message when Excel refuses an entry.
	std::wstring RuleText (const Plan::Col &c)
		{
		auto num = [] (double v) { return Csv::Tidy (v); };
		const bool lo = !std::isnan (c.lo), hi = !std::isnan (c.hi);
		switch (c.type)
			{
			case Plan::Type::Bool:
				return L"1 = on, 0 = off.";
			case Plan::Type::Double:
			case Plan::Type::Long:
				{
				const std::wstring what = c.type == Plan::Type::Long ? L"A whole number" : L"A number";
				if (lo && hi)
					return what + L" from " + num (c.lo) + L" to " + num (c.hi) + L".";
				if (lo)
					return what + L", " + num (c.lo) + L" or more.";
				if (hi)
					return what + L", " + num (c.hi) + L" or less.";
				return what + L".";
				}
			case Plan::Type::Text:
				if (!c.choices.empty ())
					return L"Pick from the list.";
				if (!std::isnan (c.hi))
					return L"Text, up to " + num (c.hi) + L" characters.";
				return L"Text.";
			}
		return std::wstring ();
		}


	/// THE MOST-USED PARAMETERS FIRST, right after the identity columns: feeds
	/// and speeds, depth of cut, stock to leave, coolant. Every other column
	/// follows in table order. Only the sheet order changes - a load finds
	/// columns by name.
	struct Front
		{
		const wchar_t *name;
		int group;
		};
	const Front kFront[] = {
		{ L"feed", GFeeds }, { L"feed_mode", GFeeds }, { L"speed", GFeeds },
		{ L"speed_mode", GFeeds }, { L"spindle_dir", GFeeds }, { L"max_ss", GFeeds },
		{ L"plunge", GFeeds }, { L"plunge_mode", GFeeds }, { L"retract", GFeeds },
		{ L"retract_mode", GFeeds }, { L"surf_fin_feed", GFeeds },
		{ L"plunge_surf_fin", GFeeds },
		// Prime turning carries its own rough and finish feeds and speeds.
		{ L"pt_rough_speed", GFeeds }, { L"pt_rough_css", GFeeds },
		{ L"pt_rough_feed_axial", GFeeds }, { L"pt_rough_axial_type", GFeeds },
		{ L"pt_rough_feed_radial", GFeeds }, { L"pt_rough_radial_type", GFeeds },
		{ L"pt_rough_feed_from", GFeeds }, { L"pt_rough_chip", GFeeds },
		{ L"pt_rough_max_feed", GFeeds },
		{ L"pt_fin_speed", GFeeds }, { L"pt_fin_css", GFeeds },
		{ L"pt_fin_feed_axial", GFeeds }, { L"pt_fin_axial_type", GFeeds },
		{ L"pt_fin_feed_radial", GFeeds }, { L"pt_fin_radial_type", GFeeds },
		{ L"pt_fin_feed_from", GFeeds }, { L"pt_fin_chip", GFeeds },
		{ L"pt_fin_max_feed", GFeeds },
		// Face / groove finish-pass feed and speed.
		{ L"use_finish_feed", GFeeds }, { L"finish_feed", GFeeds },
		{ L"use_finish_ss", GFeeds }, { L"finish_ss", GFeeds }, { L"finish_ss_css", GFeeds },
		{ L"retract_rapid", GFeeds }, { L"retract_feedrate", GFeeds },
		{ L"back_feed", GFeeds }, { L"fr_override_on", GFeeds }, { L"fr_override", GFeeds },
		{ L"ss_override_on", GFeeds }, { L"ss_override", GFeeds },
		{ L"step", GDepth }, { L"min_step", GDepth }, { L"stepover", GDepth },
		{ L"stepover_percent", GDepth }, { L"n_cuts", GDepth }, { L"peck1", GDepth },
		{ L"peck2", GDepth },
		{ L"rough_step", GDepth }, { L"rough_n_steps", GDepth }, { L"rough_step_by", GDepth },
		{ L"rough_step_percent", GDepth }, { L"finish_step", GDepth },
		{ L"dcuts_on", GDepth }, { L"dcut_rough", GDepth }, { L"dcut_finish", GDepth },
		{ L"dcut_fin_n", GDepth }, { L"dcut_stock", GDepth }, { L"mcuts_on", GDepth },
		{ L"mcut_rough_n", GDepth }, { L"mcut_rough_amt", GDepth }, { L"mcut_fin_n", GDepth },
		{ L"mcut_fin_amt", GDepth },
		{ L"stock_x", GStock }, { L"stock_z", GStock }, { L"fin_stock_x", GStock },
		{ L"fin_stock_z", GStock },
		{ L"coolant_before", GCoolant }, { L"coolant_with", GCoolant },
		{ L"coolant_after", GCoolant }, { L"coolant", GCoolant } };

	/// The group of every table column, decided from the FULL sheet order before
	/// unused columns are dropped - so dropping a group's first column cannot
	/// relabel the group.
	int GroupOf (const std::wstring &name)
		{
		static const std::map<std::wstring, int> groups = [] ()
			{
			// The first column of each group in the tables' own order. Anything the
			// front list does not take stays in the group it starts in.
			static const std::pair<const wchar_t *, int> starts[] = {
				{ L"feed", GFeeds }, { L"home_mode", GToolpath + 1 },
				{ L"ref_pt_on", GToolpath + 2 }, { L"clearance_on", GToolpath + 3 },
				{ L"filter_on", GToolpath + 4 }, { L"tplane_id", GToolpath + 5 },
				{ L"insp_do_stop", GToolpath + 6 } };
			std::map<std::wstring, int> g;
			int cur = GToolpath;
			for (const std::wstring &n : Lathe::SheetColumns ())
				{
				for (const auto &st : starts)
					if (n == st.first)
						cur = st.second;
				g[n] = cur;
				}
			for (const Front &f : kFront)
				g[f.name] = f.group;
			return g;
			} ();
		const auto it = groups.find (name);
		return it == groups.end () ? 1 : it->second;
		}

	/// A name for an operation type this tool does not read yet - what the
	/// dump reports it skipped, so the next kind to add is the one parts use.
	std::wstring OpcodeName (long code)
		{
		static const std::map<long, const wchar_t *> names = {
			{ TP_LGROOVE, L"GROOVE" }, { TP_LTHREAD, L"THREAD" }, { TP_LPOINT, L"POINT" },
			{ TP_LFACE, L"FACE" }, { TP_LCUTOFF, L"CUTOFF" }, { TP_LPLUNGE_ROUGH, L"PLUNGE ROUGH" },
			{ TP_LATHE_CONTOUR_ROUGH, L"CONTOUR ROUGH" }, { TP_LATHE_PRIME_TURNING, L"PRIME TURNING" },
			{ TP_LCAN_FINISH, L"CANNED FINISH" }, { TP_LCAN_ROUGH, L"CANNED ROUGH" },
			{ TP_LCAN_PATTERN, L"CANNED PATTERN" }, { TP_LCAN_GROOVE, L"CANNED GROOVE" },
			{ TP_LCAN_GROOVE_FINISH, L"CANNED GROOVE FINISH" }, { TP_LSTOCK_XFER, L"STOCK TRANSFER" },
			{ TP_LSTOCK_FLIP, L"STOCK FLIP" }, { TP_LBARFEED, L"BAR FEED" },
			{ TP_LCHUCK_CLAMP, L"CHUCK" }, { TP_LTAILSTOCK, L"TAILSTOCK" },
			{ TP_LSTEADYREST, L"STEADY REST" }, { TP_LPLUNGE_TURN_CHAIN, L"PLUNGE TURN" },
			{ TP_LPLUNGE_TURN_PT, L"PLUNGE TURN (POINT)" }, { TP_LPARK_TURRET, L"PARK TURRET" },
			{ TP_LATHE_THREAD_CUSTOM, L"CUSTOM THREAD" },
			{ TP_LATHE_BAXIS_CONTOUR_TURNING, L"B-AXIS CONTOUR TURN" },
			{ TP_LATHE_AAXIS_CONTOUR_TURNING, L"A-AXIS CONTOUR TURN" },
			{ TP_MANUAL_ENTRY, L"MANUAL ENTRY (MILL)" },
			{ TP_CONTOUR, L"MILL CONTOUR" }, { TP_DRILL, L"MILL DRILL" },
			{ TP_XFORM, L"TRANSFORM" }, { TP_2D_HMM, L"MILL DYNAMIC / PEEL (2D HMM)" } };
		const auto it = names.find (code);
		return it != names.end () ? std::wstring (it->second)
								  : L"opcode " + std::to_wstring (code);
		}

	/// An embedded resource's bytes (the compiled macros, the ribbon), or "".
	std::string Resource (int id)
		{
		HMODULE m = AfxGetResourceHandle ();
		HRSRC r = FindResourceW (m, MAKEINTRESOURCEW (id), RT_RCDATA);
		if (r == nullptr)
			return std::string ();
		HGLOBAL g = LoadResource (m, r);
		const DWORD n = SizeofResource (m, r);
		const char *p = g != nullptr ? static_cast<const char *> (LockResource (g)) : nullptr;
		return p != nullptr ? std::string (p, n) : std::string ();
		}

	/// One operation the dump will write.
	struct Found
		{
		operation *op = nullptr;
		const Lathe::Table *t = nullptr;
		void *prm = nullptr;
		};

	bool IsXCoolant (const std::wstring &name)
		{
		return name == L"coolant_before" || name == L"coolant_with" || name == L"coolant_after";
		}

	/// Whether a column means anything for this row. Coolant depends on the
	/// operation's MACHINE: a V9 machine has the one `coolant` setting; an X-style
	/// machine has before / with / after - shown on a V9 machine only to clear
	/// leftover X-style entries it would ignore.
	bool Applies (const Found &f, const std::wstring &name)
		{
		if (Lathe::IndexOf (*f.t, name) < 0)
			return false;
		if (name == L"coolant")
			return Coolant::IsV9 (*f.op);
		if (IsXCoolant (name))
			return !Coolant::IsV9 (*f.op) || Coolant::HasXEntries (*f.op);
		return true;
		}

	/// Whether this kind of operation carries the operation-level set (tool,
	/// feeds, planes). Manual entry does not.
	bool HasTool (const Lathe::Table &t)
		{
		return t.inspectStart > t.opLevelStart;
		}

	std::wstring Counts (const std::map<std::wstring, int> &m)
		{
		std::wstring s;
		for (const auto &kv : m)
			s += L"\r\n  " + kv.first + L":  " + std::to_wstring (kv.second);
		return s;
		}

	/// The formatted sheet for the same rows.
	/// The length at each feed, longest first: "0.01/rev 312.4; 400/min 55.2 ..." -
	/// a dynamic mill's back feed shows as its own entry.
	std::wstring FeedGroups (const Paths::Totals &t)
		{
		std::map<double, double> lenAt;
		for (const Estimate::Group &g : t.groups)
			for (const Estimate::Band &b : g.bands)
				lenAt[g.feed] += b.len;
		std::vector<std::pair<double, double>> v (lenAt.begin (), lenAt.end ());
		std::sort (v.begin (), v.end (), [] (const std::pair<double, double> &a,
											 const std::pair<double, double> &b)
				   { return a.second > b.second; });
		std::wstring s;
		for (const auto &kv : v)
			s += (s.empty () ? L"" : L"; ") + Csv::Tidy (std::fabs (kv.first))
				 + (kv.first < 0 ? L"/rev " : L"/min ")
				 + Csv::Tidy (std::round (kv.second * 1000.0) / 1000.0);
		return s;
		}

	/// The column that sets a roughing op's depth of cut (or stepover), whose
	/// change scales its cutting time; nullptr for kinds that do not rough.
	const wchar_t *DepthColumn (const std::wstring &type)
		{
		if (type == L"ROUGH" || type == L"PRIME")
			return L"step";
		if (type == L"DYNAMIC" || type == L"DYNAMIC MILL")
			return L"stepover";
		if (type == L"FACE" || type == L"GROOVE" || type == L"PLUNGE ROUGH" || type == L"CONTOUR ROUGH")
			return L"rough_step";
		return nullptr;
		}

	/// A sheet cell's time in SECONDS, whatever was typed: "9:00", "1:02:30", 540,
	/// or a value Excel turned into a time of day.
	std::wstring SecondsOf (const std::wstring &c)
		{
		return L"IF(ISNUMBER(" + c + L"),IF(" + c + L"<1," + c + L"*86400," + c + L"),IFERROR(IF(LEN(" + c
			   + L")-LEN(SUBSTITUTE(" + c + L",\":\",\"\"))=1,TIMEVALUE(\"0:\"&" + c + L"),TIMEVALUE(" + c
			   + L"))*86400,VALUE(" + c + L")))";
		}

	/// The first ISO insert code in a name ("CNMG432 BORING BAR" -> CNMG432,
	/// "RH RNG86E" -> RNG86E, "RPGV45" -> RPGV45), or "". Shape letter, a
	/// clearance letter, then letters and digits: four letters then digits, or a
	/// round's three (RNG86). A holder code (CRGNL8, PCLNL2020) is not one.
	std::wstring IsoInsertCode (const std::wstring &name)
		{
		std::wstring u;
		for (wchar_t c : name)
			u += std::iswalnum (c) ? static_cast<wchar_t> (std::towupper (c)) : L' ';
		size_t at = 0;
		while (at < u.size ())
			{
			while (at < u.size () && u[at] == L' ')
				++at;
			size_t end = at;
			while (end < u.size () && u[end] != L' ')
				++end;
			const std::wstring tok = u.substr (at, end - at);
			at = end;
			if (tok.size () < 4 || std::wstring (L"CDEKRSTVW").find (tok[0]) == std::wstring::npos
				|| std::wstring (L"ABCDEFGNP").find (tok[1]) == std::wstring::npos || !std::iswalpha (tok[2]))
				continue;
			const bool four = std::iswalpha (tok[3]) && (tok.size () == 4 || std::iswdigit (tok[4]));
			const bool round3 = tok[0] == L'R' && std::iswdigit (tok[3]);
			if (four || round3)
				return tok;
			}
		return std::wstring ();
		}

	/// An ANSI insert shape code's name and nose angle (0: round / unknown).
	std::pair<std::wstring, double> ShapeName (wchar_t code)
		{
		switch (code)
			{
			case L'C': return { L"C 80°", 80 };
			case L'D': return { L"D 55°", 55 };
			case L'E': return { L"E 75°", 75 };
			case L'K': return { L"K 55°", 55 };
			case L'M': return { L"M 86°", 86 };
			case L'V': return { L"V 35°", 35 };
			case L'T': return { L"T 60°", 60 };
			case L'S': return { L"S 90°", 90 };
			case L'W': return { L"W 80°", 80 };
			case L'H': return { L"H 120°", 120 };
			case L'O': return { L"O 135°", 135 };
			case L'P': return { L"P 108°", 108 };
			case L'R': return { L"Round", 0 };
			default: return { std::wstring (), 0 };
			}
		}

	std::wstring Dim (double v)
		{
		return Csv::Tidy (std::round (v * 1000.0) / 1000.0);
		}

	/// The insert as its geometry says - what tools are grouped by, whatever they
	/// are called: "Round IC 1", "C 80° IC 0.5 r0.031", "groove w0.197 r0.016".
	std::wstring InsertLabel (const ToolPictures::InsertInfo &in)
		{
		if (!in.ok)
			return std::wstring ();
		const std::wstring name = ShapeName (in.shape).first;
		if (in.shape == 0 && in.ic <= 0 && in.width <= 0)
			return std::wstring ();						// an empty record: the outline decides
		if (in.shape == L'R')
			return L"Round IC " + Dim (in.ic);
		if (!name.empty () && in.ic > 0)
			return name + L" IC " + Dim (in.ic) + (in.radius > 0 ? L" r" + Dim (in.radius) : L"");
		if (in.width > 0)
			return L"groove w" + Dim (in.width) + (in.radius > 0 ? L" r" + Dim (in.radius) : L"");
		return std::wstring ();
		}

	/// The ISO shape whose nose angle is nearest a measured one (within 6°), or 0.
	wchar_t NearestShape (double angle)
		{
		wchar_t best = 0;
		double off = 6.0;
		for (const wchar_t c : std::wstring (L"CDEMVTSWHOP"))
			{
			const double a = ShapeName (c).second;
			if (a > 0 && std::fabs (a - angle) <= off)
				{
				off = std::fabs (a - angle);
				best = c;
				}
			}
		return best;
		}

	/// The insert as its cutting OUTLINE says (a 3D tool, no insert data): "Round
	/// 1.069", "C 80° r0.031", else "polygon 91° r0.016".
	std::wstring OutlineLabel (const StockSim::ToolShape &sim)
		{
		if (!sim.ok)
			return std::wstring ();
		if (sim.round)
			return L"Round " + Dim (sim.size);
		const wchar_t c = NearestShape (sim.noseAngle);
		const std::wstring r = sim.noseRadius > 0 ? L" r" + Dim (sim.noseRadius) : L"";
		return (c ? ShapeName (c).first : L"polygon " + Dim (std::round (sim.noseAngle)) + L"°") + r;
		}

	/// The insert's own data against the shape the stock simulation sweeps; with
	/// no insert data, an ISO code in the names against it.
	std::wstring ShapeCheck (const ToolPictures::InsertInfo &in, const StockSim::ToolShape &sim,
							 const std::wstring &code)
		{
		if (!in.ok || (in.shape == 0 && in.ic <= 0 && in.width <= 0))		// none, or an empty record (a probe)
			{
			const std::wstring iso = IsoInsertCode (code);
			if (!sim.ok)
				return L"no insert data, no outline";
			if (iso.empty ())
				return L"no insert data - grouped by its cutting outline";
			ToolPictures::InsertInfo fromCode;
			fromCode.ok = true;
			fromCode.shape = iso[0];
			const std::wstring c = ShapeCheck (fromCode, sim, std::wstring ());
			return c == L"ok" ? L"ok (outline matches " + iso + L")" : c + L" (by the code " + iso + L")";
			}
		if (!sim.ok)
			return L"simulation has no outline - nose circle only";
		const auto [name, angle] = ShapeName (in.shape);
		if (in.shape == L'R')
			{
			if (!sim.round)
				return L"MISMATCH: insert is round, simulation sweeps a polygon";
			return in.ic > 0 && std::fabs (sim.size - in.ic) > 0.1 * in.ic
					   ? L"size differs: simulation " + Dim (sim.size) + L" dia, insert IC " + Dim (in.ic)
					   : L"ok";
			}
		if (sim.round)
			return L"MISMATCH: insert is " + (name.empty () ? std::wstring (L"not round") : name)
				   + L", simulation sweeps a round";
		if (angle > 0)
			return std::fabs (sim.noseAngle - angle) <= 8
					   ? L"ok"
					   : L"MISMATCH: nose angle " + Dim (std::round (sim.noseAngle)) + L"° in the simulation, "
							 + Dim (angle) + L"° insert";
		return L"not checked (shape " + std::wstring (in.shape ? std::wstring (1, in.shape) : L"none") + L")";
		}

	std::wstring SimShapeText (const StockSim::ToolShape &sim)
		{
		if (!sim.ok)
			return L"nose circle only";
		return sim.round ? L"round " + Dim (sim.size) + L" dia"
						 : L"polygon, nose " + Dim (std::round (sim.noseAngle)) + L"°, " + Dim (sim.size) + L" across";
		}

	/// Edges per insert from its ISO code: corners by the shape letter - C, D,
	/// E, K, V 2; T, W 3; S 4; a round (R) about 8 - doubled when the clearance
	/// letter is N (double-sided). Blank without a code: typed in instead.
	std::wstring EdgesGuess (const std::wstring &insert)
		{
		std::wstring up;
		for (wchar_t c : insert)
			up += static_cast<wchar_t> (std::towupper (c));
		if (up.find (L"ROUND") != std::wstring::npos)
			return L"8";
		std::wstring code = IsoInsertCode (insert);
		if (code.empty () && !insert.empty () && ShapeName (insert[0]).second > 0 && insert.size () > 1
			&& insert[1] == L' ')
			code = std::wstring (1, insert[0]) + L"X";		// a geometry label: shape only, one-sided
		if (insert.rfind (L"Round", 0) == 0)
			return L"8";
		if (code.empty ())
			return std::wstring ();
		int corners = 0;
		switch (code[0])
			{
			case L'C': case L'D': case L'E': case L'K': case L'V': corners = 2; break;
			case L'T': case L'W': corners = 3; break;
			case L'S': corners = 4; break;
			case L'R': return L"8";
			default: return std::wstring ();
			}
		return std::to_wstring (code[1] == L'N' ? corners * 2 : corners);
		}

	bool WriteXlsx (const std::filesystem::path &file, const std::vector<Csv::Row> &out,
					const std::vector<Column> &columns, const std::vector<Found> &rows,
					const std::vector<Stats> &stats, std::vector<Xlsx::Sheet::ToolRow> tools,
					const std::vector<int> &toolOfRow, bool macros)
		{
		Xlsx::Sheet s;
		s.rows = out;
		s.groupNames = kGroupNames;
		s.frozenCols = 5;				// op_idn, type, tool, comment, changes
		s.kindCol = 1;
		s.trackChanges = true;
		for (size_t c = 0; c < columns.size (); ++c)
			if (columns[c].name == L"changes")
				s.changesCol = static_cast<int> (c);
		s.outlineGroup.assign (kGroupNames.size (), 0);
		s.collapseGroup.assign (kGroupNames.size (), 0);
		for (int g : kOutlined)
			s.outlineGroup[static_cast<size_t> (g)] = 1;
		for (int g : kCollapsed)
			s.collapseGroup[static_cast<size_t> (g)] = 1;

		for (const Column &c : columns)
			{
			s.group.push_back (c.group);
			s.readOnly.push_back (c.readOnly);
			s.text.push_back (c.text);
			s.textFormat.push_back (c.name == L"insp_time");	// "9:00" stays "9:00"
			}

		for (const Found &f : rows)
			{
			std::vector<char> na;
			for (const Column &c : columns)
				na.push_back (!c.info && !Applies (f, c.name));
			s.notApplicable.push_back (na);
			}

		// ---- Linked pairs: the percent is a FORMULA of the amount and the tool
		// radius. The formula keeps the stored percent while the amount is
		// untouched (a compare against the dumped amount), so a sheet nobody
		// edited loads as no change; edit the amount and the percent follows.
		// Editing a percent directly replaces its formula.
		static const wchar_t *const pairs[][2] = { { L"stepover", L"stepover_percent" },
												   { L"radius", L"radius_percent" } };
		auto colOf = [&columns] (const std::wstring &n) -> int
			{
			for (size_t c = 0; c < columns.size (); ++c)
				if (columns[c].name == n)
					return static_cast<int> (c);
			return -1;
			};
		auto letters = [] (size_t col)
			{
			const std::string l = Xlsx::ColName (col);
			return std::wstring (l.begin (), l.end ());
			};

		s.formula.assign (rows.size (), std::vector<std::wstring> (columns.size ()));
		const int radCol = colOf (L"tool_radius");
		for (size_t d = 0; radCol >= 0 && d < rows.size (); ++d)
			{
			const std::vector<std::wstring> &row = s.rows[d + 1];
			if (!Xlsx::IsPlainNumber (row[static_cast<size_t> (radCol)]))
				continue;
			const std::wstring rad = L"$" + letters (static_cast<size_t> (radCol))
									 + std::to_wstring (d + 3);
			for (const auto &pr : pairs)
				{
				const int a = colOf (pr[0]), p = colOf (pr[1]);
				if (a < 0 || p < 0 || Lathe::IndexOf (*rows[d].t, pr[1]) < 0)
					continue;
				const std::wstring &amount = row[static_cast<size_t> (a)];
				const std::wstring &percent = row[static_cast<size_t> (p)];
				if (!Xlsx::IsPlainNumber (amount) || !Xlsx::IsPlainNumber (percent))
					continue;
				const std::wstring cell = letters (static_cast<size_t> (a)) + std::to_wstring (d + 3);
				s.formula[d][static_cast<size_t> (p)] =
					L"IF(" + cell + L"=" + amount + L"," + percent + L"," + cell + L"/" + rad
					+ L"*100)";
				}
			}

		// The model's feed seconds at the dumped values, per row with a live estimate
		// (NaN: none) - what the insert-flip estimate scales from.
		std::vector<double> modelOf (rows.size (), std::nan (""));

		// ---- THE LIVE ESTIMATE. Per operation: Mastercam's own time, minus what
		// the model says at the dumped values, plus the model at the sheet's
		// CURRENT values - so it starts equal to Mastercam's figure and moves
		// only with edits. Each group's feed and speed are tied to the cells they
		// came from, matched by value.
		{
		const int estCol = colOf (L"est_seconds");
		const int estText = colOf (L"est_cycle_time");
		const int changeCol = colOf (L"time_change");
		const int rawCol = colOf (L"cycle_time_raw");
		s.untracked.assign (columns.size (), 0);
		for (int c : { estCol, estText, changeCol, colOf (L"flips_est"), colOf (L"cut_seconds_est"),
					   colOf (L"mrr"), colOf (L"mrr_avg"), colOf (L"mrr_engaged"), colOf (L"flips_uncommented"),
					   colOf (L"flips_part"), colOf (L"edge_limit"), colOf (L"edge_after"), colOf (L"est_note") })
			if (c >= 0)
				s.untracked[static_cast<size_t> (c)] = 1;

		for (size_t d = 0; estCol >= 0 && d < rows.size () && d < stats.size (); ++d)
			{
			const Stats &st = stats[d];
			if (!st.path.ok || st.seconds <= 0 || st.path.groups.empty ())
				continue;
			const Lathe::Table &t = *rows[d].t;
			const std::vector<std::wstring> &row = s.rows[d + 1];
			const std::wstring rowNo = std::to_wstring (d + 3);
			auto has = [&] (const wchar_t *name)
				{ return colOf (name) >= 0 && Lathe::IndexOf (t, name) >= 0; };
			auto ref = [&] (const wchar_t *name)
				{ return letters (static_cast<size_t> (colOf (name))) + rowNo; };
			auto text = [&] (const wchar_t *name)
				{ return has (name) ? row[static_cast<size_t> (colOf (name))] : std::wstring (); };
			auto num = [&] (const wchar_t *name, double &v)
				{ return has (name) && Csv::ParseDouble (text (name), v); };
			auto on = [&] (const wchar_t *flagName)
				{
				double v = 1;
				return flagName == nullptr || !has (flagName) || !num (flagName, v) || v != 0;
				};
			auto same = [] (double a, double b)
				{ return std::fabs (a - b) <= 1e-9 * (std::max) (1.0, std::fabs (a)); };

			// Feed cells, most specific first. A pass that is switched off
			// cannot be what drove its moves.
			// axis: 1 = only for moves along Z, 2 = only along X, 0 = any.
			struct FeedCell { const wchar_t *col, *mode, *pass; int axis; };
			static const FeedCell feeds[] = {
				{ L"fr_override", nullptr, L"fr_override_on", 0 },
				{ L"pt_rough_feed_axial", L"pt_rough_axial_type", L"do_rough", 1 },
				{ L"pt_rough_feed_radial", L"pt_rough_radial_type", L"do_rough", 2 },
				{ L"pt_fin_feed_axial", L"pt_fin_axial_type", L"do_finish", 1 },
				{ L"pt_fin_feed_radial", L"pt_fin_radial_type", L"do_finish", 2 },
				{ L"finish_feed", nullptr, L"use_finish_feed", 0 },
				{ L"back_feed", nullptr, nullptr, 0 },
				{ L"retract_feedrate", nullptr, nullptr, 0 },
				{ L"feed", L"feed_mode", nullptr, 0 },
				{ L"plunge", L"plunge_mode", nullptr, 0 },
				{ L"retract", L"retract_mode", nullptr, 0 } };
			struct SpeedCell { const wchar_t *col, *css, *pass; bool word; };
			static const SpeedCell speeds[] = {
				{ L"pt_rough_speed", L"pt_rough_css", L"do_rough", false },
				{ L"pt_fin_speed", L"pt_fin_css", L"do_finish", false },
				{ L"finish_ss", L"finish_ss_css", L"use_finish_ss", false },
				{ L"ss_override", nullptr, L"ss_override_on", false },
				{ L"speed", L"speed_mode", nullptr, true } };

			std::vector<Estimate::Group> groups = st.path.groups;
			std::wstring formula;
			double model = 0;
			for (int attempt = 0; attempt < 3; ++attempt)
				{
				model = 0;
				std::wstring terms;
				for (const Estimate::Group &g : groups)
					{
					Estimate::Cells cells;
					for (const FeedCell &fc : feeds)
						{
						double v = 0;
						if (!on (fc.pass) || !num (fc.col, v) || !same (v, std::fabs (g.feed)))
							continue;
						if (fc.axis != 0 && g.axis != 0 && fc.axis != g.axis)
							continue;			// prime turning: axial feed for Z moves, radial for X
						if (fc.mode != nullptr && has (fc.mode)
							&& (text (fc.mode) == L"per rev") != (g.feed < 0))
							continue;
						cells.feed = ref (fc.col);
						if (fc.mode != nullptr && has (fc.mode) && text (fc.mode) != L"surface finish")
							cells.feedMode = ref (fc.mode);
						break;
						}
					for (const SpeedCell &sc : speeds)
						{
						double v = 0;
						if (!on (sc.pass) || !num (sc.col, v) || !same (v, g.speed))
							continue;
						bool css = false;
						if (sc.css == nullptr)
							css = false;				// an RPM override
						else if (sc.word)
							css = text (sc.css) == L"CSS";
						else
							{
							double f = 0;
							css = num (sc.css, f) && f != 0;
							}
						if (css != g.css)
							continue;
						cells.speed = ref (sc.col);
						if (sc.css != nullptr && has (sc.css))
							{
							cells.css = ref (sc.css);
							cells.cssIsWord = sc.word;
							}
						double cap = 0;
						if (num (L"max_ss", cap) && same (cap, g.cap))
							cells.cap = ref (L"max_ss");
						break;
						}
					model += Estimate::Seconds (g);
					terms += L"+" + Estimate::Term (g, cells);
					}
				formula = Csv::Tidy (st.seconds - model) + terms;
				if (formula.size () < 7800)
					break;
				// Too long for one Excel formula: fewer, wider bands.
				for (Estimate::Group &g : groups)
					{
					std::vector<Estimate::Band> merged;
					for (size_t b = 0; b < g.bands.size (); b += 2)
						{
						Estimate::Band m = g.bands[b];
						if (b + 1 < g.bands.size ())
							{
							m.len += g.bands[b + 1].len;
							m.lenDia += g.bands[b + 1].lenDia;
							}
						merged.push_back (m);
						}
					g.bands = merged;
					}
				formula.clear ();
				}
			if (formula.empty () || formula.size () >= 7800)
				continue;

			// DEPTH OF CUT: the sheet cannot change the path, but a roughing op's
			// cutting time goes close to inversely with its depth (twice the step,
			// about half the passes) - so a depth edit scales the model's part of the
			// estimate by old / new depth, until the op is regenerated.
			const wchar_t *depth = DepthColumn (t.schema.type);
			double depth0 = 0;
			if (depth != nullptr && has (depth) && num (depth, depth0) && depth0 > 0)
				{
				const std::wstring at = ref (depth);
				const size_t cut = formula.find (L'+', 1);		// after Mastercam's part (which may be negative)
				if (cut != std::wstring::npos)
					formula = formula.substr (0, cut) + L"+(" + formula.substr (cut + 1) + L")*IF(N(" + at + L")>0,"
							  + Csv::FormatDouble (depth0) + L"/" + at + L",1)";
				const int noteCol = colOf (L"est_note");
				if (noteCol >= 0)
					s.formula[d][static_cast<size_t> (noteCol)] =
						L"IF(N(" + at + L")<>" + Csv::FormatDouble (depth0)
						+ L",\"" + std::wstring (depth) + L" changed - estimate until the op is regenerated\",\"\")";
				}

			const std::wstring est = letters (static_cast<size_t> (estCol)) + rowNo;
			s.formula[d][static_cast<size_t> (estCol)] = formula;
			modelOf[d] = model;
			if (estText >= 0)
				s.formula[d][static_cast<size_t> (estText)] =
					L"IF(ISNUMBER(" + est + L"),TEXT(" + est + L"/86400,\"[h]:mm:ss\"),\"\")";
			if (changeCol >= 0 && rawCol >= 0)
				{
				const std::wstring raw = letters (static_cast<size_t> (rawCol)) + rowNo;
				s.formula[d][static_cast<size_t> (changeCol)] =
					L"IF(ISNUMBER(" + est + L"),IF(ROUND(" + est + L"-" + raw + L",0)<0,\"-\",\"+\")&TEXT(ABS(ROUND("
					+ est + L"-" + raw + L",0))/86400,\"[h]:mm:ss\"),\"\")";
				}
			}
		}

		// ---- INSERT FLIPS, LIVE. Feed time now = the live estimate less Mastercam's
		// figure plus the model at the dumped values - the model at the sheet's
		// CURRENT values. Each cause of a flip scales on its own, calibrated so the
		// dumped values give the NCI's own count exactly:
		//   time      INT(feed now / interval * c), c = interval0 * (n + 0.5) / feed0 -
		//             the real interval runs over the setting (stops wait for a gap
		//             between cuts), and c carries that over; n + 0.5 keeps the
		//             count from flipping on a hair's change.
		//   distance  the same on cut length, which feeds and speeds do not change.
		//   end, cuts as dumped; the end stop follows its switch.
		// Time is on but nothing stopped (a long cut with no gap)? It stays at 0.
		{
		const int flipsEst = colOf (L"flips_est");
		const int cutEst = colOf (L"cut_seconds_est");
		const int rawCol = colOf (L"cycle_time_raw");
		const int estCol = colOf (L"est_seconds");
		auto full = [] (double v)
			{
			wchar_t buf[40];
			swprintf_s (buf, L"%.12g", v);
			return std::wstring (buf);
			};
		for (size_t d = 0; d < rows.size () && d < stats.size (); ++d)
			{
			const Stats &st = stats[d];
			if (!st.path.ok)
				continue;
			const Lathe::Table &t = *rows[d].t;
			const std::wstring rowNo = std::to_wstring (d + 3);
			auto has = [&] (const wchar_t *name)
				{ return colOf (name) >= 0 && Lathe::IndexOf (t, name) >= 0; };
			auto ref = [&] (const wchar_t *name)
				{ return letters (static_cast<size_t> (colOf (name))) + rowNo; };

			const bool live = !std::isnan (modelOf[d]) && rawCol >= 0 && estCol >= 0;
			const double feed0 = live ? modelOf[d] : st.path.feedSeconds;
			const std::wstring feedNow = live ? L"(" + letters (static_cast<size_t> (estCol)) + rowNo + L"-"
													+ letters (static_cast<size_t> (rawCol)) + rowNo + L"+"
													+ full (modelOf[d]) + L")"
											  : full (feed0);
			if (cutEst >= 0 && live)
				{
				s.formula[d][static_cast<size_t> (cutEst)] = feedNow.substr (1, feedNow.size () - 2);
				s.rows[d + 1][static_cast<size_t> (cutEst)] = Csv::Tidy (feed0);
				}

			const Inspect::Settings &is = st.insp;
			const Inspect::Result &r = st.inspRes;
			if (flipsEst < 0 || !st.inspOk || !is.doStop || !has (L"insp_do_stop"))
				continue;
			if (r.flips == 0 && !Inspect::IsFlip (is.comment))
				continue;						// its stops are not flips: 0 whatever changes

			std::wstring terms = std::to_wstring (r.other);
			if (has (L"insp_time_on") && has (L"insp_time") && feed0 > 0)
				{
				double c = 1;
				if (r.byTime > 0)
					c = is.time * (r.byTime + 0.5) / feed0;
				else if (is.timeOn && is.time > 0 && feed0 >= is.time)
					c = 0;
				if (c > 0)
					terms += L"+IF(" + ref (L"insp_time_on") + L"=1,IFERROR(INT(" + feedNow + L"/"
							 + SecondsOf (ref (L"insp_time")) + L"*" + full (c) + L"),0),0)";
				}
			if (has (L"insp_dist_on") && has (L"insp_dist") && st.path.cutLength > 0)
				{
				const double len = st.path.cutLength;
				double c = 1;
				if (r.byDist > 0)
					c = is.dist * (r.byDist + 0.5) / len;
				else if (is.distOn && is.dist > 0 && len >= is.dist)
					c = 0;
				if (c > 0)
					terms += L"+IF(" + ref (L"insp_dist_on") + L"=1,IFERROR(INT(" + full (len) + L"/"
							 + ref (L"insp_dist") + L"*" + full (c) + L"),0),0)";
				}
			if (has (L"insp_at_end"))
				{
				const int e = r.atEnd > 0 ? 1 : (is.atEnd ? 0 : 1);
				if (e > 0)
					terms += L"+IF(" + ref (L"insp_at_end") + L"=1,1,0)";
				}
			s.formula[d][static_cast<size_t> (flipsEst)] = L"IF(" + ref (L"insp_do_stop") + L"=1," + terms + L",0)";
			}

		// ---- MATERIAL REMOVAL, THEORETICAL AND LIVE. While cutting:
		//   turning   12 x SFM x feed per rev x depth of cut   (in^3/min)
		//             Vc x f x ap                              (cm^3/min, metric)
		//   drilling  pi D^2 / 4 x feed per minute
		// at the op's mean cutting diameter (from the NCI): an RPM op's surface
		// speed is taken there, and CSS is capped by max_ss there. Over the whole
		// op, leads, rapids and air included: mrr x cut time / op time. Removed =
		// mrr x cut time - an estimate (it counts air cuts as cutting).
		{
		const int mrrCol = colOf (L"mrr");
		const int basisCol = colOf (L"mrr_basis"), diaCol = colOf (L"cut_dia");
		for (size_t d = 0; mrrCol >= 0 && d < rows.size () && d < stats.size (); ++d)
			{
			const Stats &st = stats[d];
			if (!st.path.ok || st.path.cutLength <= 0)
				continue;
			const Lathe::Table &t = *rows[d].t;
			const std::vector<std::wstring> &row = s.rows[d + 1];
			const std::wstring rowNo = std::to_wstring (d + 3);
			auto has = [&] (const wchar_t *name)
				{ return colOf (name) >= 0 && Lathe::IndexOf (t, name) >= 0; };
			auto ref = [&] (const wchar_t *name)
				{ return letters (static_cast<size_t> (colOf (name))) + rowNo; };
			auto text = [&] (const wchar_t *name)
				{ return colOf (name) >= 0 ? row[static_cast<size_t> (colOf (name))] : std::wstring (); };
			auto num = [&] (const wchar_t *name)
				{
				double v = 0;
				Csv::ParseDouble (text (name), v);
				return v;
				};
			if (!has (L"speed") || !has (L"feed") || !has (L"feed_mode") || !has (L"speed_mode"))
				continue;
			const std::wstring type = t.schema.type;
			const bool drill = type == L"DRILL" || type == L"MILL DRILL";
			const wchar_t *ap = nullptr;
			if (type == L"ROUGH" || type == L"PRIME")
				ap = L"step";
			else if (type == L"FINISH" && num (L"n_cuts") > 1)
				ap = L"step";						// a finish's step is a depth only between several passes
			else if (type == L"DYNAMIC")
				ap = L"stepover";
			else if (has (L"rough_step"))
				ap = L"rough_step";					// face, groove, plunge rough
			if (!drill && (ap == nullptr || !has (ap) || num (ap) <= 0))
				continue;
			const bool mm = text (L"units") == L"mm";
			const double kSpeed = mm ? 1000.0 : 12.0;		// surface speed <-> RPM
			const double pi = 3.14159265358979323846;

			// The cached numbers: the same arithmetic as the formulas.
			// Rounded as the cut_dia cell shows it - the formula reads the cell.
			const double dia = drill ? st.toolDia
									 : std::round (2.0 * st.path.lengthTimesRadius / st.path.cutLength * 1000.0) / 1000.0;
			if (!(dia > 0))
				continue;
			const bool css = text (L"speed_mode") == L"CSS";
			const double speed = num (L"speed"), cap = has (L"max_ss") ? num (L"max_ss") : 0;
			double rpm = css ? speed * kSpeed / (pi * dia) : speed;
			if (css && cap > 0)
				rpm = (std::min) (rpm, cap);
			const double sfm = rpm * pi * dia / kSpeed;
			const bool perRev = text (L"feed_mode") == L"per rev";
			const double feed = num (L"feed");
			const double ipr = perRev ? feed : (rpm > 0 ? feed / rpm : 0);
			double mrr = drill ? pi * dia * dia / 4.0 * ipr * rpm / (mm ? 1000.0 : 1.0)
							   : (mm ? 1.0 : 12.0) * sfm * ipr * num (ap);
			mrr = std::round (mrr * 1000.0) / 1000.0;

			const std::wstring K = mm ? L"1000" : L"12";
			const std::wstring D = drill ? Csv::Tidy (dia) : ref (L"cut_dia");
			const std::wstring capRef = has (L"max_ss") ? ref (L"max_ss") : L"0";
			const std::wstring rpmF = L"IF(" + ref (L"speed_mode") + L"=\"CSS\",IF(N(" + capRef + L")>0,MIN(" + ref (L"speed")
									  + L"*" + K + L"/(PI()*" + D + L")," + capRef + L")," + ref (L"speed") + L"*" + K
									  + L"/(PI()*" + D + L"))," + ref (L"speed") + L")";
			const std::wstring iprF = L"IF(" + ref (L"feed_mode") + L"=\"per rev\"," + ref (L"feed") + L"," + ref (L"feed")
									  + L"/" + rpmF + L")";
			const std::wstring mrrF = drill
				? L"ROUND(PI()*" + D + L"^2/4*" + iprF + L"*" + rpmF + (mm ? L"/1000" : L"") + L",3)"
				: L"ROUND(" + std::wstring (mm ? L"1" : L"12") + L"*" + rpmF + L"*PI()*" + D + L"/" + K + L"*" + iprF + L"*"
					  + ref (ap) + L",3)";
			s.formula[d][static_cast<size_t> (mrrCol)] = L"IFERROR(" + mrrF + L",\"\")";
			s.rows[d + 1][static_cast<size_t> (mrrCol)] = Csv::Tidy (mrr);
			if (drill && diaCol >= 0)
				s.rows[d + 1][static_cast<size_t> (diaCol)] = Csv::Tidy (dia);

			if (basisCol >= 0)
				s.rows[d + 1][static_cast<size_t> (basisCol)] =
					drill ? L"drill " + Csv::Tidy (dia) + L" dia x feed per minute"
						  : std::wstring (ap) + L" " + text (ap) + L" x feed per rev x surface speed at "
							+ Csv::Tidy (std::round (dia * 1000.0) / 1000.0) + L" dia"
							+ (css && cap > 0 && speed * kSpeed / (pi * dia) > cap ? L" (capped by max_ss)" : L"")
							+ (mm ? L" - cm3/min" : L" - in3/min");
			}
		}

		// ---- WHAT EACH OP REMOVES, AND HOW FAST. The volume is the toolpath's: it
		// does not change with the sheet's feeds and speeds (a depth-of-cut edit
		// changes it only once the op is regenerated), so it is a fixed number -
		// from the 2D stock simulation where that covers the op, else mrr x cut
		// time at the dumped values. The RATES are live, from that volume and the
		// live times:
		//   mrr_avg      removed / op time            (rapids, leads and air in it)
		//   mrr_engaged  removed / cut time not in air - set it beside mrr
		{
		const int remCol = colOf (L"removed"), fromCol = colOf (L"removed_from"), airCol = colOf (L"air_pct");
		const int avgCol = colOf (L"mrr_avg"), engCol = colOf (L"mrr_engaged"), mrrCol = colOf (L"mrr");
		const int cutEst = colOf (L"cut_seconds_est"), estCol = colOf (L"est_seconds"), rawCol = colOf (L"cycle_time_raw");
		for (size_t d = 0; remCol >= 0 && cutEst >= 0 && estCol >= 0 && rawCol >= 0 && d < rows.size () && d < stats.size (); ++d)
			{
			const Stats &st = stats[d];
			if (!st.path.ok)
				continue;
			std::vector<std::wstring> &row = s.rows[d + 1];
			const std::wstring rowNo = std::to_wstring (d + 3);
			auto cell = [&] (int c) { return letters (static_cast<size_t> (c)) + rowNo; };
			double cut0 = 0, mrr0 = 0;
			Csv::ParseDouble (row[static_cast<size_t> (cutEst)], cut0);
			double removed = 0, air = 0;
			std::wstring from;
			if (st.simOk)
				{
				removed = st.simRemoved;
				// The air share is the simulation's: shown only where its volume
				// agrees with Mastercam's (or there is nothing to set it against).
				air = st.simHasAir && st.simAgrees ? st.simAirPct : 0;
				from = st.simFromBoundary ? L"Mastercam's stock boundaries"
										  : st.simBoundaryBad ? L"stock simulation (Mastercam's boundary looks wrong)"
															  : L"stock simulation";
				}
			else if (mrrCol >= 0 && Csv::ParseDouble (row[static_cast<size_t> (mrrCol)], mrr0) && mrr0 > 0)
				{
				removed = mrr0 * cut0 / 60.0;
				from = L"estimate: mrr x cut time (counts air as cutting)";
				}
			else
				continue;
			removed = std::round (removed * 1000.0) / 1000.0;
			row[static_cast<size_t> (remCol)] = Csv::Tidy (removed);
			if (fromCol >= 0)
				row[static_cast<size_t> (fromCol)] = from;
			if (const int chk = colOf (L"removed_check"); chk >= 0)
				row[static_cast<size_t> (chk)] = st.simCheck;
			if (airCol >= 0 && st.simOk && st.simHasAir && st.simAgrees)
				row[static_cast<size_t> (airCol)] = Csv::Tidy (std::round (air * 10.0) / 10.0);

			const bool live = !s.formula[d][static_cast<size_t> (estCol)].empty ();
			const std::wstring tot = cell (live ? estCol : rawCol);
			const double tot0 = st.seconds;
			if (avgCol >= 0 && tot0 > 0)
				{
				s.formula[d][static_cast<size_t> (avgCol)] = L"IFERROR(ROUND(" + cell (remCol) + L"/" + tot + L"*60,3),\"\")";
				row[static_cast<size_t> (avgCol)] = Csv::Tidy (std::round (removed / tot0 * 60.0 * 1000.0) / 1000.0);
				}
			if (engCol >= 0 && st.simOk && st.simHasAir && st.simAgrees && cut0 > 0 && air < 100)
				{
				s.formula[d][static_cast<size_t> (engCol)] = L"IFERROR(ROUND(" + cell (remCol) + L"/(" + cell (cutEst) + L"*(1-"
															 + cell (airCol) + L"/100))*60,3),\"\")";
				const double air1 = std::round (air * 10.0) / 10.0;
				row[static_cast<size_t> (engCol)] = Csv::Tidy (std::round (removed / (cut0 * (1 - air1 / 100)) * 60.0 * 1000.0) / 1000.0);
				}
			}
		}

		// ---- THE EDGE CLOCK. Per tool, in program order, the cut time an edge has
		// done since its last flip - tool changes in between do not reset it.
		//   commented stops (CHANGE/ROTATE INSERT) are flips, as counted above; the
		//     clock leaves the op at the cut time after its last flip.
		//   UNCOMMENTED stops leave nothing in the NCI, so they are inferred from
		//     the settings: one at each edge_limit of cut while the clock runs
		//     (the timer on), or one at the end (the end stop) - each a flip only
		//     once the edge has cut edge_limit. Short ops let the time carry on.
		//   edge_limit = the op's insp_time when one is filled in (even with the
		//     timer off), else the tool's edge life on the Tools page (8:00).
		// Each op reads the clock from the previous op of its tool BY op_idn, so
		// sorting or filtering the sheet does not break the chain.
		const int toolsN = static_cast<int> (tools.size ());
		{
		const int idCol = colOf (L"op_idn"), toolCol = colOf (L"tool");
		const int uncCol = colOf (L"flips_uncommented"), partCol = colOf (L"flips_part");
		const int limCol = colOf (L"edge_limit"), afterCol = colOf (L"edge_after");
		const int cutCol = colOf (L"cut_seconds_est"), estFlips = colOf (L"flips_est");
		auto full = [] (double v)
			{
			wchar_t buf[40];
			swprintf_s (buf, L"%.12g", v);
			return std::wstring (buf);
			};
		const std::wstring lastRow = std::to_wstring (rows.size () + 2);
		auto range = [&] (int c)
			{
			const std::wstring l = letters (static_cast<size_t> (c));
			return L"$" + l + L"$3:$" + l + L"$" + lastRow;
			};
		const double defaultLife = 480.0;		// the Tools page's edge life to start with
		// The Tools page's edge life for a tool (column H, as m:ss), in seconds.
		const std::wstring lifeOf = toolsN > 0
			? L"IFERROR(" + SecondsOf (L"VLOOKUP(" + std::wstring (L"TOOLREF") + L"&\"\",'Tools'!$A$2:$H$"
									   + std::to_wstring (toolsN + 1) + L",8,FALSE)") + L"," + full (defaultLife) + L")"
			: full (defaultLife);
		std::map<long, double> clockOf;			// tool -> edge clock leaving its last op
		std::map<long, std::wstring> lastOpOf;	// tool -> op_idn of its last op
		for (size_t d = 0; idCol >= 0 && toolCol >= 0 && uncCol >= 0 && partCol >= 0 && limCol >= 0
						   && afterCol >= 0 && cutCol >= 0 && d < rows.size () && d < stats.size (); ++d)
			{
			const Stats &st = stats[d];
			if (!st.path.ok || !HasTool (*rows[d].t))
				continue;
			const Lathe::Table &t = *rows[d].t;
			const std::wstring rowNo = std::to_wstring (d + 3);
			auto has = [&] (const wchar_t *name)
				{ return colOf (name) >= 0 && Lathe::IndexOf (t, name) >= 0; };
			auto ref = [&] (const wchar_t *name)
				{ return letters (static_cast<size_t> (colOf (name))) + rowNo; };
			auto cell = [&] (int c) { return letters (static_cast<size_t> (c)) + rowNo; };
			const long tool = rows[d].op->tl.tlno;
			std::vector<std::wstring> &row = s.rows[d + 1];

			// The clock coming in.
			const double in0 = clockOf.count (tool) ? clockOf[tool] : 0.0;
			const std::wstring in = lastOpOf.count (tool)
				// INDIRECT, not INDEX over the column: a range holding this very cell
				// would be a circular reference.
				? L"IFERROR(INDIRECT(\"" + letters (static_cast<size_t> (afterCol)) + L"\"&(MATCH(" + lastOpOf[tool] + L","
				  + range (idCol) + L",0)+2)),0)"
				: std::wstring (L"0");
			double feed0 = 0;
			Csv::ParseDouble (row[static_cast<size_t> (cutCol)], feed0);
			const std::wstring feed = cell (cutCol);

			// The limit.
			const Inspect::Settings &is = st.insp;
			const bool hasTime = has (L"insp_time");
			const double lim0 = hasTime && is.time > 0 ? is.time : defaultLife;
			std::wstring life = lifeOf;
			for (size_t at = life.find (L"TOOLREF"); at != std::wstring::npos; at = life.find (L"TOOLREF"))
				life.replace (at, 7, cell (toolCol));
			s.formula[d][static_cast<size_t> (limCol)] = hasTime
				? L"IF(" + SecondsOf (ref (L"insp_time")) + L">0," + SecondsOf (ref (L"insp_time")) + L"," + life + L")"
				: life;
			row[static_cast<size_t> (limCol)] = Csv::Tidy (lim0);
			const std::wstring lim = cell (limCol);

			// Flips the comment-less stops add, and the clock going out.
			const bool on = st.inspOk && is.doStop;
			double unc0 = 0, out0 = in0 + feed0;
			if (on && is.commentOn)
				{
				if (st.inspRes.flips > 0 && st.path.feedSeconds > 0)
					out0 = feed0 * st.inspRes.tail / st.path.feedSeconds;
				}
			else if (on && (is.timeOn || is.atEnd))
				{
				const double total = in0 + feed0;
				unc0 = std::floor (total / lim0);
				if (!is.timeOn)
					unc0 = (std::min) (unc0, 1.0);
				out0 = is.atEnd && unc0 > 0 ? 0.0 : total - unc0 * lim0;
				}
			if (st.inspOk && has (L"insp_do_stop") && has (L"insp_comment_on") && has (L"insp_at_end"))
				{
				const std::wstring doStop = ref (L"insp_do_stop"), comment = ref (L"insp_comment_on");
				const std::wstring timeOn = has (L"insp_time_on") ? ref (L"insp_time_on") : L"0";
				const std::wstring atEnd = ref (L"insp_at_end"), total = L"(" + in + L"+" + feed + L")";
				const std::wstring n = L"INT(" + total + L"/" + lim + L")";
				s.formula[d][static_cast<size_t> (uncCol)] = L"IF(AND(" + doStop + L"=1," + comment + L"<>1),IF(" + timeOn
					+ L"=1," + n + L",IF(" + atEnd + L"=1,MIN(1," + n + L"),0)),0)";
				const std::wstring unc = cell (uncCol);
				const std::wstring tailRatio = st.inspRes.flips > 0 && st.path.feedSeconds > 0
					? full (st.inspRes.tail / st.path.feedSeconds) : std::wstring ();
				const std::wstring commented = tailRatio.empty () || estFlips < 0
					? total
					: L"IF(N(" + cell (estFlips) + L")>0," + feed + L"*" + tailRatio + L"," + total + L")";
				s.formula[d][static_cast<size_t> (afterCol)] = L"IF(" + doStop + L"<>1," + total + L",IF(" + comment
					+ L"=1," + commented + L",IF(AND(" + atEnd + L"=1," + unc + L">0),0," + total + L"-" + unc + L"*"
					+ lim + L")))";
				}
			else
				s.formula[d][static_cast<size_t> (afterCol)] = in + L"+" + feed;
			row[static_cast<size_t> (uncCol)] = Csv::Tidy (unc0);
			row[static_cast<size_t> (afterCol)] = Csv::Tidy (std::round (out0 * 1000.0) / 1000.0);

			// All the flips of the op: commented (as counted) + inferred.
			const double flips0 = st.inspRes.flips + unc0;
			s.formula[d][static_cast<size_t> (partCol)] = (estFlips >= 0 ? L"N(" + cell (estFlips) + L")" : std::wstring (L"0"))
														  + L"+" + cell (uncCol);
			row[static_cast<size_t> (partCol)] = Csv::Tidy (flips0);

			clockOf[tool] = out0;
			lastOpOf[tool] = row[static_cast<size_t> (idCol)];
			}
		}

		// ---- THE TOOLS PAGE: per tool, its insert and - live - flips and cut time
		// per part, summed from the main sheet by tool number; then per insert.
		const int toolCol = colOf (L"tool");
		const int partCol = colOf (L"flips_part");
		if (!tools.empty () && toolCol >= 0 && partCol >= 0 && cutEst >= 0)
			{
			const std::wstring last = std::to_wstring (rows.size () + 2);
			auto range = [&] (int c)
				{
				const std::wstring l = letters (static_cast<size_t> (c));
				return L"'Lathe params'!$" + l + L"$3:$" + l + L"$" + last;
				};
			s.toolExtraHeads = { L"Insert", L"Flips / part", L"Cut time / part", L"Longest between flips",
								 L"Edge life (fallback)", L"Insert code", L"Simulated shape", L"Shape check" };
			std::vector<double> flipsOf (tools.size (), 0), cutOf (tools.size (), 0), longOf (tools.size (), 0);
			std::vector<char> inspects (tools.size (), 0);	// any of its ops has tool inspection on
			std::vector<std::map<double, int>> lifeVotes (tools.size ());	// its inspected ops' insp_time
			for (size_t d = 0; d < rows.size () && d < stats.size (); ++d)
				if (d < toolOfRow.size () && toolOfRow[d] >= 0)
					{
					const size_t k = static_cast<size_t> (toolOfRow[d]);
					double fp = 0;
					Csv::ParseDouble (s.rows[d + 1][static_cast<size_t> (partCol)], fp);
					flipsOf[k] += fp;
					double v = 0;
					if (Csv::ParseDouble (s.rows[d + 1][static_cast<size_t> (cutEst)], v))
						cutOf[k] += v;
					longOf[k] = (std::max) (longOf[k], stats[d].inspRes.longest);
					inspects[k] = inspects[k] || (stats[d].inspOk && stats[d].insp.doStop);
					if (stats[d].inspOk && stats[d].insp.doStop && stats[d].insp.time > 0)
						++lifeVotes[k][stats[d].insp.time];
					}
			std::map<std::wstring, std::vector<size_t>> byInsert;
			for (size_t k = 0; k < tools.size (); ++k)
				{
				Xlsx::Sheet::ToolRow &tr = tools[k];
				// extra[0] the insert (what tools group by), then code, simulated
				// shape and check - kept for the end.
				const std::vector<Xlsx::Sheet::FreeCell> tail (tr.extra.size () > 1 ? tr.extra.begin () + 1 : tr.extra.end (),
															   tr.extra.end ());
				tr.extra.resize (1);
				tr.extra[0].editable = true;			// the insert: correct it, and the table below follows
				const std::wstring row = std::to_wstring (k + 2);
				Xlsx::Sheet::FreeCell f, c, l;
				f.formula = L"SUMIF(" + range (toolCol) + L",$A" + row + L"," + range (partCol) + L")";
				f.text = Csv::Tidy (flipsOf[k]);
				c.formula = L"TEXT(SUMIF(" + range (toolCol) + L",$A" + row + L"," + range (cutEst)
							+ L")/86400,\"[h]:mm:ss\")";
				c.text = Hms (cutOf[k]);
				l.text = longOf[k] > 0 ? Inspect::MinSec (longOf[k]) : std::wstring ();
				tr.extra.push_back (f);
				tr.extra.push_back (c);
				tr.extra.push_back (l);
				// Edge life (typed, m:ss): the cut time before a comment-less stop
				// counts as a flip, on this tool's ops that set no insp_time of their
				// own - the flips themselves come from the inspection criteria. It
				// starts at the tool's own insp_time (the most common across its
				// inspected ops, the shorter on a tie), 8:00 where none is set, and
				// blank for a tool that does not inspect (its insert lasts the part).
				Xlsx::Sheet::FreeCell life;
				double life0 = 480.0;
				int best = 0;
				for (const auto &v : lifeVotes[k])			// ascending: the shorter wins a tie
					if (v.second > best)
						{
						best = v.second;
						life0 = v.first;
						}
				life.text = inspects[k] ? Inspect::MinSec (life0) : L"";
				life.editable = true;
				life.textFormat = true;
				tr.extra.push_back (life);
				for (const Xlsx::Sheet::FreeCell &t : tail)
					tr.extra.push_back (t);
				if (!tr.extra[0].text.empty ())
					byInsert[tr.extra[0].text].push_back (k);
				}
			{
				// Below the tools: per insert. The tool rows' column D is the insert
				// and E the flips, so each insert sums its tools by name - type a name
				// into a tool's Insert cell and into a row here, and it adds up.
				const std::wstring lastTool = std::to_wstring (tools.size () + 1);
				auto head = [] (const wchar_t *t)
					{
					Xlsx::Sheet::FreeCell h;
					h.text = t;
					h.head = true;
					return h;
					};
				s.toolsAfter.push_back ({ head (L"Inserts"), head (L"Insert"), head (L"Used by"),
										  head (L"Edges per insert"), head (L"Flips / part"), head (L"Inserts / part") });
				size_t i = 0;
				for (const auto &kv : byInsert)
					{
					const std::wstring r = std::to_wstring (tools.size () + 4 + i);
					Xlsx::Sheet::FreeCell blank, name, used, edges, flips, inserts;
					name.text = kv.first;
					name.editable = true;
					double total = 0;
					for (size_t k : kv.second)
						{
						used.text += (used.text.empty () ? L"T" : L", T") + tools[k].number;
						total += flipsOf[k];
						}
					// Edges: from an ISO code on one of its tools (CNMG432 -> 4), else
					// from the shape alone (one-sided - type over a double-sided one).
					std::wstring codeOf;
					for (size_t k : kv.second)
						if (codeOf.empty () && tools[k].extra.size () > 5)
							codeOf = tools[k].extra[5].text;
					edges.text = EdgesGuess (codeOf);
					if (edges.text.empty ())
						edges.text = EdgesGuess (kv.first);
					edges.editable = true;
					flips.formula = L"SUMIF($D$2:$D$" + lastTool + L",$B" + r + L",$E$2:$E$" + lastTool + L")";
					flips.text = Csv::Tidy (total);
					double e = 0;
					inserts.formula = L"IF(N($D" + r + L")>0,ROUNDUP(ROUND($E" + r + L"/$D" + r + L",6),0),\"\")";
					inserts.text = Csv::ParseDouble (edges.text, e) && e > 0
									   ? Csv::Tidy (std::ceil (std::round (total / e * 1e6) / 1e6)) : std::wstring ();
					s.toolsAfter.push_back ({ blank, name, used, edges, flips, inserts });
					++i;
					}
				// Room for the inserts the names did not give.
				for (int extra = 0; extra < 4; ++extra, ++i)
					{
					const std::wstring r = std::to_wstring (tools.size () + 4 + i);
					Xlsx::Sheet::FreeCell blank, name, used, edges, flips, inserts;
					name.editable = true;
					edges.editable = true;
					flips.formula = L"IF($B" + r + L"=\"\",\"\",SUMIF($D$2:$D$" + lastTool + L",$B" + r + L",$E$2:$E$"
									+ lastTool + L"))";
					inserts.formula = L"IF(N($D" + r + L")>0,ROUNDUP(ROUND(N($E" + r + L")/$D" + r + L",6),0),\"\")";
					s.toolsAfter.push_back ({ blank, name, used, edges, flips, inserts });
					}
				}

			// PART TOTALS, as dumped and now: the effect of the edits on the whole
			// part. "As dumped" sums the hidden Dumped sheet; "Now" the live sheet.
			const int estSecCol = colOf (L"est_seconds"), remCol = colOf (L"removed");
			if (estSecCol >= 0)
				{
				auto head = [] (const wchar_t *t)
					{
					Xlsx::Sheet::FreeCell h;
					h.text = t;
					h.head = true;
					return h;
					};
				auto sumOf = [&] (int c, bool dumped)
					{
					const std::wstring l = letters (static_cast<size_t> (c));
					return L"SUM(" + std::wstring (dumped ? L"Dumped!" : L"'Lathe params'!") + L"$" + l + L"$3:$" + l
						   + L"$" + last + L")";
					};
				double est0 = 0, cut0 = 0, flips0 = 0, rem0 = 0;
				for (size_t d = 0; d < rows.size (); ++d)
					{
					double v = 0;
					if (Csv::ParseDouble (s.rows[d + 1][static_cast<size_t> (estSecCol)], v)) est0 += v;
					if (Csv::ParseDouble (s.rows[d + 1][static_cast<size_t> (cutEst)], v)) cut0 += v;
					if (Csv::ParseDouble (s.rows[d + 1][static_cast<size_t> (partCol)], v)) flips0 += v;
					if (remCol >= 0 && Csv::ParseDouble (s.rows[d + 1][static_cast<size_t> (remCol)], v)) rem0 += v;
					}
				auto timeRow = [&] (const wchar_t *what, int c, double v0)
					{
					Xlsx::Sheet::FreeCell a, label, was, now, change;
					label.text = what;
					const std::wstring w = sumOf (c, true), n = sumOf (c, false);
					was.formula = L"TEXT(" + w + L"/86400,\"[h]:mm:ss\")";
					was.text = Hms (v0);
					now.formula = L"TEXT(" + n + L"/86400,\"[h]:mm:ss\")";
					now.text = Hms (v0);
					change.formula = L"IF(ROUND(" + n + L"-" + w + L",0)<0,\"-\",\"+\")&TEXT(ABS(ROUND(" + n + L"-" + w
									 + L",0))/86400,\"[h]:mm:ss\")";
					change.text = L"+0:00:00";
					return std::vector<Xlsx::Sheet::FreeCell> { a, label, was, now, change };
					};
				auto numRow = [&] (const wchar_t *what, const std::wstring &w, const std::wstring &n, double v0)
					{
					Xlsx::Sheet::FreeCell a, label, was, now, change;
					label.text = what;
					was.formula = L"ROUND(" + w + L",2)";
					was.text = Csv::Tidy (std::round (v0 * 100.0) / 100.0);
					now.formula = L"ROUND(" + n + L",2)";
					now.text = was.text;
					change.formula = L"ROUND(" + n + L"-(" + w + L"),2)";
					change.text = L"0";
					return std::vector<Xlsx::Sheet::FreeCell> { a, label, was, now, change };
					};
				s.toolsAfter.push_back ({});
				s.toolsAfter.push_back ({ head (L"Part totals"), head (L"What"), head (L"As dumped"), head (L"Now"),
										  head (L"Change") });
				s.toolsAfter.push_back (timeRow (L"Cycle time (estimate)", estSecCol, est0));
				s.toolsAfter.push_back (timeRow (L"Cutting time", cutEst, cut0));
				s.toolsAfter.push_back (numRow (L"Insert flips per part", sumOf (partCol, true), sumOf (partCol, false), flips0));
				if (remCol >= 0)
					{
					s.toolsAfter.push_back (numRow (L"Material removed (fixed by the toolpaths)", sumOf (remCol, true),
													sumOf (remCol, false), rem0));
					const double mrr0 = est0 > 0 ? rem0 / est0 * 60.0 : 0;
					s.toolsAfter.push_back (numRow (L"Average removal rate over the cycle (per min)",
													L"IFERROR(" + sumOf (remCol, true) + L"/" + sumOf (estSecCol, true) + L"*60,0)",
													L"IFERROR(" + sumOf (remCol, false) + L"/" + sumOf (estSecCol, false) + L"*60,0)",
													mrr0));
					}
				}
			}
		}
		s.tools = tools;

		// ---- WHAT EVERY CELL ACCEPTS, checked by Excel as it is typed, with a
		// tooltip saying what the column is. The rules are the load's own limits,
		// so a value Excel lets in is one the load will take. Cells that a load
		// never writes - read-only, or not this kind of operation's - refuse any
		// typing and say why. Cells sharing one rule share one entry.
		std::map<std::wstring, Xlsx::Sheet::Validation> rules;
		std::map<std::wstring, std::vector<std::pair<size_t, size_t>>> cellsOf;
		auto put = [&] (const Xlsx::Sheet::Validation &v, size_t col, size_t row)
			{
			std::wstring key = std::wstring (v.type.begin (), v.type.end ()) + L"|"
							   + std::wstring (v.op.begin (), v.op.end ()) + L"|" + v.f1 + L"|"
							   + v.f2 + L"|" + v.title + L"|" + v.prompt + L"|" + v.error;
			for (const std::wstring &c : v.choices)
				key += L"|" + c;
			if (!rules.count (key))
				rules[key] = v;
			cellsOf[key].push_back ({ col, row });
			};

		for (size_t c = 0; c < columns.size (); ++c)
			{
			const Column &col = columns[c];
			const std::wstring help = ColumnHelp::For (col.name);
			const Plan::Col *pc = col.info ? nullptr : PlanCol (col.name);
			for (size_t d = 0; d < rows.size (); ++d)
				{
				Xlsx::Sheet::Validation v;
				v.title = col.name;
				const bool applies = col.info || Applies (rows[d], col.name);
				const bool kindHasIt = col.info || Lathe::IndexOf (*rows[d].t, col.name) >= 0;
				if (!applies)
					{
					v.type = "custom";
					v.f1 = L"FALSE";
					if (!kindHasIt)
						v.prompt = L"Does not apply to " + rows[d].t->schema.type + L" operations - leave it blank.";
					else if (col.name == L"coolant")
						v.prompt = L"This operation's machine uses X-style coolant - see coolant_before / with / after.";
					else
						v.prompt = L"This operation's machine uses V9 coolant - see the coolant column.";
					v.error = v.prompt;
					}
				else if (col.info || col.readOnly || pc == nullptr || pc->readOnly)
					{
					v.type = "custom";
					v.f1 = L"FALSE";
					v.prompt = (help.empty () ? L"" : help + L"\n") + L"Read-only: shown for reference, never written back.";
					v.error = L"This column is read-only - it is never written back to Mastercam.";
					}
				else
					{
					const std::wstring rule = RuleText (*pc);
					v.prompt = (help.empty () ? L"" : help + L"\n") + rule + L" Blank = leave as is.";
					v.error = rule;
					const bool lo = !std::isnan (pc->lo), hi = !std::isnan (pc->hi);
					if (IsXCoolant (col.name) && Coolant::IsV9 (*rows[d].op))
						{
						// Leftovers on a V9 machine: the only thing to do is clear them.
						v.type = "list";
						v.choices = { L"none" };
						v.prompt = L"This operation's machine uses V9 coolant, so these X-style entries are ignored. "
								   L"Set none to clear them; the coolant column is what the machine uses.";
						v.error = L"Only none - this machine uses V9 coolant (the coolant column).";
						}
					else if (IsXCoolant (col.name))
						{
						v.type = "list";
						v.choices = Coolant::Choices (*rows[d].op);
						}
					else if (!pc->choices.empty ())
						{
						v.type = "list";
						v.choices = pc->choices;
						}
					else if (pc->type == Plan::Type::Bool)
						{
						v.type = "list";
						v.choices = { L"0", L"1" };
						}
					else if (pc->type == Plan::Type::Double || pc->type == Plan::Type::Long)
						{
						const bool whole = pc->type == Plan::Type::Long;
						v.type = whole ? "whole" : "decimal";
						const std::wstring dLo = whole ? L"-2147483648" : L"-1E+100";
						const std::wstring dHi = whole ? L"2147483647" : L"1E+100";
						if (lo && hi)
							{ v.op = "between"; v.f1 = Csv::Tidy (pc->lo); v.f2 = Csv::Tidy (pc->hi); }
						else if (lo)
							{ v.op = "greaterThanOrEqual"; v.f1 = Csv::Tidy (pc->lo); }
						else if (hi)
							{ v.op = "lessThanOrEqual"; v.f1 = Csv::Tidy (pc->hi); }
						else
							{ v.op = "between"; v.f1 = dLo; v.f2 = dHi; }
						}
					else if (pc->type == Plan::Type::Text && hi)
						{
						v.type = "textLength";
						v.op = "lessThanOrEqual";
						v.f1 = Csv::Tidy (pc->hi);
						}
					}
				put (v, c, d + 3);
				}
			}

		// Cells to compact ranges: a run of rows in one column is one range.
		for (auto &kv : rules)
			{
			std::vector<std::pair<size_t, size_t>> &cells = cellsOf[kv.first];
			std::sort (cells.begin (), cells.end ());
			std::string ref;
			for (size_t k = 0; k < cells.size ();)
				{
				size_t e = k;
				while (e + 1 < cells.size () && cells[e + 1].first == cells[k].first
					   && cells[e + 1].second == cells[e].second + 1)
					++e;
				const std::string col = Xlsx::ColName (cells[k].first);
				ref += (ref.empty () ? "" : " ") + col + std::to_string (cells[k].second);
				if (e > k)
					ref += ":" + col + std::to_string (cells[e].second);
				k = e + 1;
				}
			kv.second.cells = ref;
			s.validations.push_back (kv.second);
			}

		// ---- With macros: the compiled VBA and the ribbon tab (.xlsm).
		if (macros)
			{
			s.vbaProject = Resource (IDR_VBAPROJECT);
			s.ribbonXml = Resource (IDR_RIBBON);
			}

		// ---- The Tools page, and each row's tool number linking to it.
		s.tools = tools;
		const int toolCol = colOf (L"tool");
		for (size_t d = 0; toolCol >= 0 && d < rows.size () && d < toolOfRow.size (); ++d)
			if (toolOfRow[d] >= 0)
				s.toolLinks.push_back ({ Xlsx::ColName (static_cast<size_t> (toolCol)) + std::to_string (d + 3),
										 static_cast<size_t> (toolOfRow[d]) });

		return Xlsx::Write (file, s);
		}
	}

namespace Dump
	{
	int Run ()
		{
		Coolant::Reset ();
		// EVERY DUMP GETS ITS OWN FILES: the date and time are in the name, so a
		// dump never overwrites an earlier one (or an edited sheet).
		wchar_t stampBuf[32] = L"";
		{
		const std::time_t now = std::time (nullptr);
		std::tm tmv = {};
		localtime_s (&tmv, &now);
		std::wcsftime (stampBuf, 32, L"%Y%m%d-%H%M%S", &tmv);
		}
		const std::wstring stamp = stampBuf;

		const std::filesystem::path part = Util::PartFile ();
		if (const unsigned long diag = Settings::Diag ())
			Util::Log (part, L"diag " + std::to_wstring (diag) + L": "
							 + ((diag & 1) ? L"no Mastercam cycle time  " : L"")
							 + ((diag & 2) ? L"no NCI walk  " : L"")
							 + ((diag & 4) ? L"stock sim" : L""));
		if (part.empty ())
			{
			Util::Say (L"Save the part first - the CSV file is written beside "
					   L"it, and there is nowhere to put it yet.");
			return 0;
			}


		// ---- EVERY OPERATION THIS TOOL KNOWS, in Operation Manager order (the
		// order they run), and which of them are selected.
		Cnc::Tool::TpPartOpList &opList = TpMainOpMgr.GetMainOpList ();
		std::vector<Found> all, selected;
		std::map<std::wstring, int> allByType, selByType, skipped;
		std::vector<std::wstring> probe;
		for (INT_PTR i = 0; i < opList.GetSize (); ++i)
			{
			operation *pOp = opList.GetAt (i);
			if (pOp == nullptr)
				continue;
			const Lathe::Table *t = Lathe::TableFor (static_cast<long> (pOp->opcode));
			if (t == nullptr)
				{
				++skipped[OpcodeName (static_cast<long> (pOp->opcode))];
				probe.push_back (L"  op " + std::to_wstring (pOp->op_idn) + L"  "
								 + OpcodeName (static_cast<long> (pOp->opcode)) + L" ("
								 + std::to_wstring (static_cast<long> (pOp->opcode)) + L")  tool "
								 + std::to_wstring (pOp->tl.tlno) + L"  " + Comment (pOp)
								 + (pOp->db.nci_flag ? L"  [needs regen]" : L""));
				continue;
				}
			void *prm = Lathe::PrmFor (*pOp, t->opcode);
			if (prm == nullptr)
				continue;
			all.push_back ({ pOp, t, prm });
			++allByType[t->schema.type];
			if (pOp->db.select_flag)
				{
				selected.push_back ({ pOp, t, prm });
				++selByType[t->schema.type];
				}
			}

		// What this tool does not read yet, by kind - said in the question and logged.
		std::wstring skippedText, skippedLine;
		if (!skipped.empty ())
			{
			std::wstring line;
			for (const auto &kv : skipped)
				line += (line.empty () ? L"" : L", ") + kv.first + L" x" + std::to_wstring (kv.second);
			skippedText = L"\r\n\r\nNot read by this tool yet (left out):\r\n  " + line;
			skippedLine = line;
			Util::Log (part, L"skipped (not read yet): " + line);
			for (const std::wstring &p : probe)
				Util::Log (part, p);
			}

		if (all.empty ())
			{
			Util::Say (L"This part has no operations of a kind this tool reads, so "
					   L"there is nothing to dump." + skippedText);
			return 0;
			}

		// ---- ASK, IN A WINDOW: which operations (ticked in a tree of toolpath
		// groups - to start, what is selected in the Operation Manager, else all),
		// where it goes and what it is called. All but the ticks remembered.
		Settings::Dump settings = Settings::LoadDump ();
		std::vector<DumpDialog::Op> ops;
		for (const Found &f : all)
			{
			DumpDialog::Op o;
			o.idn = f.op->op_idn;
			o.group = f.op->cmn.grp_idn;
			o.groupName = GroupName (f.op->cmn.grp_idn);
			o.kind = f.t->schema.type;
			o.text = L"op " + std::to_wstring (f.op->op_idn) + L"    " + f.t->schema.type
					 + (HasTool (*f.t) ? L"    T" + std::to_wstring (f.op->tl.tlno) : std::wstring ())
					 + L"    " + Comment (f.op) + (f.op->db.nci_flag ? L"    [needs regen]" : L"");
			o.selectedInMgr = f.op->db.select_flag;
			o.on = selected.empty () || o.selectedInMgr;
			ops.push_back (o);
			}
		if (!DumpDialog::Show (part.wstring (), ops, skippedLine, settings))
			return 0;
		Settings::SaveDump (settings);

		std::vector<Found> chosen;
		std::map<std::wstring, int> byType;
		for (size_t i = 0; i < all.size () && i < ops.size (); ++i)
			if (ops[i].on)
				{
				chosen.push_back (all[i]);
				++byType[all[i].t->schema.type];
				}
		if (chosen.empty ())
			return 0;
		const bool onlySelected = chosen.size () < all.size ();

		// What each lathe op really removes: a 2D stock simulation of the toolpaths,
		// or - switched off - Mastercam's own stock boundary before and after each op
		// (StockSim.h). The simulation's detail goes to the log with diag bit 4.
		const std::map<long, StockSim::Result> sim =
			StockSim::Run (part, (Settings::Diag () & 4) != 0, settings.stockSim);
		const std::vector<Found> &rows = chosen;

		// ---- THE COLUMNS. Identity, then the reader's context, then only the
		// table columns some dumped operation actually has - a sheet of three
		// finish passes carries no drill or dynamic columns.
		std::vector<Column> columns;
		auto info = [&columns] (const wchar_t *name, bool text, bool isDouble, int group)
			{
			Column c;
			c.name = name;
			c.group = group;
			c.readOnly = true;
			c.text = text;
			c.isDouble = isDouble;
			c.info = true;
			columns.push_back (c);
			};
		info (L"op_idn", false, false, GIdentity);
		info (L"type", true, false, 0);
		info (L"tool", false, false, 0);
		info (L"comment", true, false, 0);
		info (L"changes", false, false, 0);
		info (L"tool_radius", false, true, 0);
		info (L"tool_name", true, false, 0);
		info (L"group_name", true, false, 0);
		info (L"units", true, false, 0);
		info (L"needs_regen", true, false, 0);

		std::vector<std::wstring> order;
		for (const Front &f : kFront)
			order.push_back (f.name);
		for (const std::wstring &name : Lathe::SheetColumns ())
			{
			bool inFront = false;
			for (const Front &f : kFront)
				inFront = inFront || name == f.name;
			if (!inFront)
				order.push_back (name);
			}

		for (const std::wstring &name : order)
			{
			Column c;
			c.name = name;
			c.group = GroupOf (name);
			bool used = false;
			for (const Found &f : rows)
				{
				const int at = Lathe::IndexOf (*f.t, name);
				if (at < 0 || !Applies (f, name))
					continue;
				used = true;
				const Plan::Col &pc = f.t->schema.cols[static_cast<size_t> (at)];
				c.readOnly = c.readOnly || pc.readOnly;
				c.text = c.text || pc.type == Plan::Type::Text;
				c.isDouble = c.isDouble || pc.type == Plan::Type::Double;
				}
			if (used)
				columns.push_back (c);
			// The stats sit after the coolant columns - whether or not `coolant` itself
			// is shown (an X-style machine has no use for it).
			if (name == L"coolant")
				{
				info (L"canned_text_raw", true, false, c.group);
				// What a person reads first in each group, then the detail; the
				// numbers the formulas work with are together, folded away.
				info (L"cycle_time", true, false, GStats);
				info (L"est_cycle_time", true, false, GStats);
				info (L"time_change", true, false, GStats);
				info (L"est_note", true, false, GStats);
				info (L"cut_length", false, true, GStats);
				info (L"rapid_length", false, true, GStats);
				info (L"travel_x_min", false, true, GStats);
				info (L"travel_x_max", false, true, GStats);
				info (L"travel_z_min", false, true, GStats);
				info (L"travel_z_max", false, true, GStats);
				info (L"feed_groups", true, false, GStats);
				info (L"flips_part", false, false, GFlips);
				info (L"flips", false, false, GFlips);
				info (L"flips_est", false, true, GFlips);
				info (L"flips_uncommented", false, false, GFlips);
				info (L"flips_why", true, false, GFlips);
				info (L"flip_longest", true, false, GFlips);
				info (L"insp_mode", true, false, GFlips);
				info (L"removed", false, true, GMrr);
				info (L"removed_from", true, false, GMrr);
				info (L"removed_check", true, false, GMrr);
				info (L"mrr", false, true, GMrr);
				info (L"mrr_engaged", false, true, GMrr);
				info (L"mrr_avg", false, true, GMrr);
				info (L"air_pct", false, true, GMrr);
				info (L"cut_dia", false, true, GMrr);
				info (L"mrr_basis", true, false, GMrr);
				info (L"cycle_time_raw", false, true, GCalc);
				info (L"est_seconds", false, true, GCalc);
				info (L"cut_seconds_est", false, true, GCalc);
				info (L"edge_limit", false, true, GCalc);
				info (L"edge_after", false, true, GCalc);
				}
			}

		// ---- THE ROWS.
		std::vector<Csv::Row> out;
		Csv::Row head;
		for (const Column &c : columns)
			head.push_back (c.name);
		out.push_back (head);

		// Two different tools sharing a number: Mastercam asks about it when it
		// works out times and reports. Say which, so the part can be fixed.
		{
		std::map<long, std::map<long, std::vector<long>>> slotsOf;	// tool no -> slot -> ops
		for (const Found &f : rows)
			if (HasTool (*f.t))
				slotsOf[f.op->tl.tlno][f.op->tl.slot].push_back (f.op->op_idn);
		for (const auto &no : slotsOf)
			if (no.second.size () > 1)
				{
				std::wstring line = L"tool number " + std::to_wstring (no.first) + L" is more than one tool:";
				for (const auto &sl : no.second)
					{
					line += L"  slot " + std::to_wstring (sl.first) + L" (op";
					for (long id : sl.second)
						line += L" " + std::to_wstring (id);
					line += L")";
					}
				Util::Log (part, line);
				}
		}

		int radiusHow[3] = { 0, 0, 0 };		// Radius::Corner, HalfDia, Guess
		std::vector<Stats> rowStats;
		for (const Found &f : rows)
			{
			operation *pOp = f.op;
			const Lathe::Table &t = *f.t;
			const bool tool = HasTool (t);

			auto read = [&] (const std::wstring &name) -> std::wstring
				{
				const int at = Lathe::IndexOf (t, name);
				return at < 0 || !Applies (f, name) ? std::wstring ()
							  : Lathe::Read (t.bindings[static_cast<size_t> (at)], pOp, f.prm);
				};

			// tool_radius: only kinds with percent-of-radius columns need it.
			std::wstring toolRadius;
			if (Lathe::IndexOf (t, L"stepover_percent") >= 0)
				{
				Radius how;
				const double r = ToolRadius (pOp, _wtof (read (L"stepover").c_str ()),
											 _wtof (read (L"stepover_percent").c_str ()), how);
				if (r > 0)
					toolRadius = Csv::FormatDouble (r);
				++radiusHow[static_cast<int> (how)];
				}

			Stats stats = tool ? StatsOf (pOp) : Stats ();
			if (tool)
				stats.toolDia = pOp->tl.dia;
			if (const auto hit = sim.find (pOp->op_idn); hit != sim.end ())
				{
				// What the op removes: MASTERCAM'S OWN stock boundaries before and
				// after it where they look sound - what Mastercam itself works out -
				// else the simulation. The two are compared: a big difference points
				// at a tool whose outline is not its insert. (The raster works in the
				// part's units: a metric part's mm^3 shown as cm^3.)
				const StockSim::Result &h = hit->second;
				const double k = pOp->tl.mm ? 1000.0 : 1.0;
				const bool hasMc = h.mcRemoved != -1;
				const bool simDone = h.simRemoved >= 0 && h.simOk;
				// A boundary that goes backwards, or jumps far past what was swept, is
				// one Mastercam has not kept up to date.
				const bool mcSound = hasMc && h.mcRemoved >= -0.001
									 && !(simDone && h.mcRemoved > 10 * h.simRemoved + 100);
				if (mcSound || (h.fromBoundary && h.ok))
					{
					stats.simOk = true;
					stats.simRemoved = (std::max) (0.0, h.mcRemoved) / k;
					stats.simFromBoundary = true;
					}
				else if (simDone)
					{
					stats.simOk = true;
					stats.simRemoved = h.simRemoved / k;
					stats.simBoundaryBad = hasMc;
					}
				if (simDone)
					{
					stats.simOwn = h.simRemoved / k;
					stats.simAirPct = h.airPct;
					stats.simHasAir = h.hasAir;
					}
				if (mcSound && simDone)
					{
					const double a = h.simRemoved / k, b = h.mcRemoved / k;
					stats.simAgrees = std::fabs (a - b) <= (std::max) (1.0, 0.2 * b);
					stats.simCheck = stats.simAgrees
						? L"simulation agrees (" + Csv::Tidy (std::round (a * 10) / 10) + L")"
						: L"simulation " + Csv::Tidy (std::round (a * 10) / 10) + L" vs Mastercam "
							  + Csv::Tidy (std::round (b * 10) / 10) + L" - check the tool's shape";
					}
				else if (stats.simBoundaryBad)
					stats.simCheck = L"Mastercam's boundary here looks wrong - simulation used";
				}
			if (tool)
				Util::Log (part, Paths::Describe (*pOp, stats.path, stats.seconds));
			if (tool && stats.path.ok)
				{
				// Tool inspection: the settings, then what the NCI did and why.
				Inspect::Settings &is = stats.insp;
				auto on = [&] (const wchar_t *name) { return read (name) == L"1"; };
				auto number = [&] (const wchar_t *name)
					{
					double v = 0;
					Csv::ParseDouble (read (name), v);
					return v;
					};
				stats.inspOk = Lathe::IndexOf (t, L"insp_do_stop") >= 0;
				if (stats.inspOk)
					{
					is.doStop = on (L"insp_do_stop");
					is.timeOn = on (L"insp_time_on");
					is.distOn = on (L"insp_dist_on");
					is.cutsOn = on (L"insp_n_cuts_on");
					is.firstCut = on (L"insp_first_cut");
					is.eachDepth = on (L"insp_each_depth");
					is.eachGroove = on (L"insp_each_groove");
					is.eachSection = on (L"insp_each_section");
					is.atEnd = on (L"insp_at_end");
					is.betweenCuts = on (L"insp_between_cuts");
					is.commentOn = on (L"insp_comment_on");
					Inspect::ParseMinSec (read (L"insp_time"), is.time);
					is.dist = number (L"insp_dist");
					is.minCut = number (L"insp_min_cut");
					is.cuts = static_cast<long> (number (L"insp_n_cuts"));
					is.sections = static_cast<long> (number (L"insp_sections"));
					is.comment = read (L"insp_comment");
					}
				else
					is.betweenCuts = true;
				stats.inspRes = Inspect::Explain (is, stats.path);
				if (is.doStop || stats.inspRes.stops > 0)
					Util::Log (part, Inspect::Describe (pOp->op_idn, is, stats.inspRes));
				}
			rowStats.push_back (stats);

			Csv::Row row;
			for (const Column &c : columns)
				{
				std::wstring v;
				if (!c.info)
					v = read (c.name);
				else if (c.name == L"op_idn")
					v = std::to_wstring (pOp->op_idn);
				else if (c.name == L"type")
					v = t.schema.type;
				else if (c.name == L"tool")
					v = std::to_wstring (pOp->tl.tlno);
				else if (c.name == L"comment")
					v = Comment (pOp);
				else if (c.name == L"tool_radius")
					v = toolRadius;
				else if (c.name == L"tool_name" && tool)
					v = std::wstring (pOp->tl.comment, wcsnlen (pOp->tl.comment, COMMENT_SIZE));
				else if (c.name == L"group_name")
					v = GroupName (pOp->cmn.grp_idn);
				else if (c.name == L"units" && tool)
					v = pOp->tl.mm ? L"mm" : L"in";
				else if (c.name == L"needs_regen")
					v = pOp->db.nci_flag ? L"yes" : L"no";
				else if (c.name == L"canned_text_raw" && tool)
					v = CannedText (pOp->cantxt);
				else if (c.name == L"changes")
					v = L"0";
				else if (c.name == L"cycle_time")
					v = stats.time;
				else if (c.name == L"cycle_time_raw")
					v = stats.timeRaw;
				else if (c.name == L"travel_x_min")
					v = stats.xMin;
				else if (c.name == L"travel_x_max")
					v = stats.xMax;
				else if (c.name == L"travel_z_min")
					v = stats.zMin;
				else if (c.name == L"travel_z_max")
					v = stats.zMax;
				else if (c.name == L"cut_length")
					v = stats.cut;
				else if (c.name == L"rapid_length")
					v = stats.rapid;
				else if (c.name == L"feed_groups" && stats.path.ok)
					v = FeedGroups (stats.path);
				else if (c.name == L"est_cycle_time" && stats.path.ok && stats.seconds > 0)
					v = stats.time;					// what the formula gives before any edit
				else if (c.name == L"time_change" && stats.path.ok && stats.seconds > 0)
					v = L"+0:00:00";
				else if (c.name == L"est_seconds" && stats.path.ok && stats.seconds > 0)
					v = stats.timeRaw;
				else if ((c.name == L"flips" || c.name == L"flips_est") && stats.path.ok)
					v = std::to_wstring (stats.inspRes.flips);
				else if (c.name == L"flips_why" && stats.path.ok)
					v = stats.inspRes.why;
				else if (c.name == L"flip_longest" && stats.inspRes.flips > 0)
					v = Inspect::MinSec (stats.inspRes.longest);
				else if (c.name == L"insp_mode" && stats.inspOk && stats.insp.doStop)
					v = stats.inspRes.mode;
				else if (c.name == L"cut_seconds_est" && stats.path.ok)
					v = Csv::Tidy (stats.path.feedSeconds);
				else if (c.name == L"cut_dia" && stats.path.ok && stats.path.cutLength > 0
						 && stats.path.lengthTimesRadius > 0
						 // A diameter means something only for turning: a mill op's X is not a radius.
						 && ((t.opcode >= TP_LATHE_START1 && t.opcode <= TP_LATHE_END1)
							 || (t.opcode >= TP_LATHE_START2 && t.opcode <= TP_LATHE_END2)))
					v = Csv::Tidy (std::round (2.0 * stats.path.lengthTimesRadius / stats.path.cutLength * 1000.0) / 1000.0);

				// Mastercam's arithmetic noise off the last digits. Well inside
				// what a load counts as the same value, so nothing untouched
				// reads as a change.
				double d = 0;
				if (c.isDouble && !v.empty () && Csv::ParseDouble (v, d))
					v = Csv::Tidy (d);
				row.push_back (v);
				}
			out.push_back (row);
			}

		// ---- ONE FILE: the workbook. It is what a person edits and what a load
		// reads straight back - no Save As CSV step, and Excel keeps every number
		// at full precision inside it (a CSV holds what the cell DISPLAYS).
		// The tool pictures, one per tool, when asked for.
		std::vector<Xlsx::Sheet::ToolRow> tools;
		std::vector<int> toolOfRow (rows.size (), -1);
			{
			std::map<long, size_t> bySlot;
			for (size_t d = 0; d < rows.size (); ++d)
				{
				const operation *o = rows[d].op;
				if (!HasTool (*rows[d].t))
					continue;
				// Lathe kinds only: asked for a mill tool's slot, the lathe lookup
				// answers with some lathe tool - a wrong picture is worse than none.
				const long code = rows[d].t->opcode;
				const bool latheKind = (code >= TP_LATHE_START1 && code <= TP_LATHE_END1)
									   || (code >= TP_LATHE_START2 && code <= TP_LATHE_END2);
				const long slot = o->tl.slot;
				size_t k = 0;
				const auto it = bySlot.find (slot);
				if (it == bySlot.end ())
					{
					Xlsx::Sheet::ToolRow tr;
					tr.number = std::to_wstring (o->tl.tlno);
					tr.name = std::wstring (o->tl.comment, wcsnlen (o->tl.comment, COMMENT_SIZE));
					std::wstring why = L"a mill tool - no picture yet";
					if (settings.pictures
						&& (!latheKind || !ToolPictures::LatheTool (slot, tr.png, tr.width, tr.height, why)))
						{
						tr.png.clear ();
						Util::Log (part, L"tool " + tr.number + L": no picture - " + why);
						}
					// The insert, for counting flips per insert (lathe tools only, as above):
					// its geometry from the tool manager is what groups tools; the names
					// (Mastercam's, else an ISO code in the tool's name) are shown beside;
					// and the shape the stock simulation sweeps is checked against it.
					Xlsx::Sheet::FreeCell ins, code, simShape, check;
					if (latheKind)
						{
						// Often the insert's 3D file ("CNMG 432.stp"): shown without the extension.
						ins.text = ToolPictures::LatheInsert (slot);
						const size_t dot = ins.text.find_last_of (L'.');
						if (dot != std::wstring::npos && ins.text.size () - dot <= 5
							&& ins.text.find_first_of (L"\\/ ", dot) == std::wstring::npos)
							ins.text.erase (dot);
						// A generic name ("Insert", nothing) says nothing: the ISO code in
						// the tool's own name, if it has one ("CNMG432 BORING BAR").
						std::wstring low;
						for (wchar_t c : ins.text)
							low += static_cast<wchar_t> (std::towlower (c));
						// A generic name: the tool's manufacturer code (a 3D tool's
						// order code), else an ISO code in the tool's own name.
						if (low.empty () || low == L"insert" || low == L"default")
							{
							std::wstring mfg = ToolPictures::LatheMfgCode (slot);
							while (!mfg.empty () && mfg.back () == L' ')
								mfg.pop_back ();
							ins.text = !mfg.empty () ? mfg : IsoInsertCode (tr.name);
							}
						const ToolPictures::InsertInfo info = ToolPictures::LatheInsertInfo (slot);
						const StockSim::ToolShape sim = StockSim::ShapeOfTool (slot);
						code.text = ins.text;
						std::wstring label = InsertLabel (info);
						if (label.empty ())
							label = OutlineLabel (sim);		// a 3D tool: what its cutting outline is
						if (!label.empty ())
							ins.text = label;
						simShape.text = SimShapeText (sim);
						check.text = ShapeCheck (info, sim, code.text);
						Util::Log (part, L"tool " + tr.number + L" insert: " + (label.empty () ? L"(no geometry)" : label)
										 + L" | code " + (code.text.empty () ? L"-" : code.text)
										 + (info.grade.empty () ? L"" : L" | grade " + info.grade)
										 + L" | simulated " + simShape.text + L" | " + check.text);
						}
					tr.extra.push_back (ins);
					tr.extra.push_back (code);
					tr.extra.push_back (simShape);
					tr.extra.push_back (check);
					k = tools.size ();
					tools.push_back (tr);
					bySlot[slot] = k;
					}
				else
					k = it->second;
				tools[k].usedBy += (tools[k].usedBy.empty () ? L"op " : L", ") + std::to_wstring (o->op_idn);
				toolOfRow[d] = static_cast<int> (k);
				}
			}

		const std::filesystem::path folder = settings.folder.empty ()
												 ? part.parent_path () : std::filesystem::path (settings.folder);
		const std::filesystem::path file = FileRules::Unique (
			folder, FileRules::Name (settings.pattern, part.stem ().wstring (), std::time (nullptr),
									 onlySelected, rows.size (), settings.macros ? L".xlsm" : L".xlsx"));
		if (!WriteXlsx (file, out, columns, rows, rowStats, tools, toolOfRow, settings.macros))
			{
			Util::Say (L"Could not write " + file.wstring () + L"\r\n\r\nCheck the "
					   L"folder can be written to, and try again.",
					   MB_ICONWARNING);
			return 0;
			}

		if (radiusHow[0] + radiusHow[1] + radiusHow[2] > 0)
			Util::Log (part, L"tool_radius: from corner radius " + std::to_wstring (radiusHow[0])
							 + L", from half diameter " + std::to_wstring (radiusHow[1])
							 + L", unmatched (guessed) " + std::to_wstring (radiusHow[2]));

		const std::wstring which = onlySelected ? L"selected" : L"all";
		Util::Log (part, L"dump: " + file.filename ().wstring () + L", "
						 + std::to_wstring (rows.size ()) + L" operation(s) (" + which + L"), "
						 + std::to_wstring (columns.size ()) + L" columns");

		Util::Say (L"Dumped " + std::to_wstring (rows.size ()) + L" operation(s) (" + which
				   + L"), in Operation Manager order:" + Counts (byType)
				   + L"\r\n\r\nto\r\n  " + file.filename ().wstring () + L"\r\nin\r\n  "
				   + file.parent_path ().wstring ()
				   + L"\r\n\r\nEdit it in Excel, SAVE it, and use \"Lathe params - load\" "
				   L"to bring the changes back.\r\n\r\n"
				   L"Leave op_idn and type alone - they are how each row finds its "
				   L"operation. Edited cells turn yellow; the changes column counts "
				   L"them. Hover a cell for what it is and what it takes.");

		Settings::SetLastDump (file.wstring (), part.wstring ());

		// Straight into Excel - the next thing anyone does with it.
		if (settings.openExcel)
			ShellExecuteW (nullptr, L"open", file.c_str (), nullptr, nullptr, SW_SHOWNORMAL);
		return 0;
		}
	}
