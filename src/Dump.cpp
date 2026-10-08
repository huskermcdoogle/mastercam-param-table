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
#include "Estimate.h"
#include "DumpDialog.h"
#include "Settings.h"
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
		L"Where it sits (read-only)", L"Tool inspection", L"Toolpath stats (read-only)" };
	enum { GIdentity, GFeeds, GDepth, GStock, GCoolant, GToolpath,
		   GHome, GRef, GPlanes, GFilter, GWhere, GInspect, GStats };

	/// Groups with a +/- of their own, and the ones that start folded away -
	/// rarely edited, still one click from view.
	const int kOutlined[] = { GFeeds, GDepth, GStock, GCoolant, GStats, GToolpath, GHome,
							  GRef, GPlanes, GFilter, GWhere, GInspect };
	const int kCollapsed[] = { GRef, GFilter, GWhere, GInspect };

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
		};

	Stats StatsOf (operation *pOp)
		{
		Stats st;
		const double t = CalcCycleTime (pOp, false);
		st.seconds = t;
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

	bool WriteXlsx (const std::filesystem::path &file, const std::vector<Csv::Row> &out,
					const std::vector<Column> &columns, const std::vector<Found> &rows,
					const std::vector<Stats> &stats, const std::vector<Xlsx::Sheet::ToolRow> &tools,
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
		for (int c : { estCol, estText, changeCol })
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
			for (int attempt = 0; attempt < 3; ++attempt)
				{
				double model = 0;
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

			const std::wstring est = letters (static_cast<size_t> (estCol)) + rowNo;
			s.formula[d][static_cast<size_t> (estCol)] = formula;
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

		// ---- ASK, IN A WINDOW: selected or all, which kinds, where it goes and
		// what it is called. All of it remembered for next time.
		Settings::Dump settings = Settings::LoadDump ();
		std::vector<DumpDialog::Kind> kinds;
		for (const auto &kv : allByType)
			{
			DumpDialog::Kind k;
			k.name = kv.first;
			k.all = kv.second;
			const auto sel = selByType.find (kv.first);
			k.selected = sel == selByType.end () ? 0 : sel->second;
			k.on = settings.skipKinds.find (L"|" + kv.first + L"|") == std::wstring::npos;
			kinds.push_back (k);
			}
		bool onlySelected = !selected.empty ();
		if (!DumpDialog::Show (part.wstring (), kinds, skippedLine, onlySelected, settings))
			return 0;
		Settings::SaveDump (settings);

		std::set<std::wstring> kindOn;
		for (const DumpDialog::Kind &k : kinds)
			if (k.on)
				kindOn.insert (k.name);
		std::vector<Found> chosen;
		std::map<std::wstring, int> byType;
		for (const Found &f : onlySelected ? selected : all)
			if (kindOn.count (f.t->schema.type))
				{
				chosen.push_back (f);
				++byType[f.t->schema.type];
				}
		if (chosen.empty ())
			return 0;
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
				info (L"cycle_time", true, false, GStats);
				info (L"cycle_time_raw", false, true, GStats);
				info (L"travel_x_min", false, true, GStats);
				info (L"travel_x_max", false, true, GStats);
				info (L"travel_z_min", false, true, GStats);
				info (L"travel_z_max", false, true, GStats);
				info (L"cut_length", false, true, GStats);
				info (L"rapid_length", false, true, GStats);
				info (L"feed_groups", true, false, GStats);
				info (L"est_cycle_time", true, false, GStats);
				info (L"time_change", true, false, GStats);
				info (L"est_seconds", false, true, GStats);
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

			const Stats stats = tool ? StatsOf (pOp) : Stats ();
			if (tool)
				Util::Log (part, Paths::Describe (*pOp, stats.path, stats.seconds));
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
		if (settings.pictures)
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
					if (!latheKind || !ToolPictures::LatheTool (slot, tr.png, tr.width, tr.height, why))
						{
						tr.png.clear ();
						Util::Log (part, L"tool " + tr.number + L": no picture - " + why);
						}
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
