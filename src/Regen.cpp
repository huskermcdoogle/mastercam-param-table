#include "stdafx.h"
#include "MastercamSdk.h"
#include "Regen.h"
#include "Settings.h"
#include "Undo.h"
#include "Util.h"

#include <algorithm>
#include <fstream>
#include <iterator>
#include <set>

namespace
	{
	/// An operation of the part by its number, as the operation list holds it
	/// now - or nullptr when it is not in the part.
	operation *Find (long id)
		{
		Cnc::Tool::TpPartOpList &ops = TpMainOpMgr.GetMainOpList ();
		for (INT_PTR i = 0; i < ops.GetSize (); ++i)
			if (operation *o = ops.GetAt (i))
				if (o->op_idn == id)
					return o;
		return nullptr;
		}

	/// Where an operation sits in the Operation Manager (the order it runs), or
	/// past the end when it is not in the part.
	INT_PTR PlaceOf (long id)
		{
		Cnc::Tool::TpPartOpList &ops = TpMainOpMgr.GetMainOpList ();
		for (INT_PTR i = 0; i < ops.GetSize (); ++i)
			if (const operation *o = ops.GetAt (i))
				if (o->op_idn == id)
					return i;
		return ops.GetSize ();
		}

	std::wstring Seconds (double s)
		{
		wchar_t buf[32];
		swprintf_s (buf, L"%.1f s", s);
		return buf;
		}

	bool ReadFile (const std::filesystem::path &file, std::wstring &text)
		{
		std::ifstream in (file, std::ios::binary);
		if (!in)
			return false;
		const std::string bytes ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		text = Csv::FromUtf8 (bytes);
		return true;
		}
	}

namespace Regen
	{
	Outcome Run (const std::filesystem::path &part, const std::vector<long> &ops,
				 const std::map<long, double> &estimates)
		{
		Outcome out;
		// Mastercam busy with its own background regeneration: starting more beside
		// it is not this tool's to do.
		const long busy = OperationInProcess (0, false);
		if (busy != 0)
			{
			out.busy = true;
			Util::Log (part, L"regen: Mastercam is regenerating already (OperationInProcess " + std::to_wstring (busy)
							 + L") - nothing started");
			return out;
			}
		Util::Log (part, L"regen begin: " + std::to_wstring (ops.size ()) + L" operation(s) asked for");
		Cnc::Tool::TpPartOpList &opList = TpMainOpMgr.GetMainOpList ();
		const ULONGLONG start = GetTickCount64 ();
		for (long id : ops)
			{
			History::Regenerated r;
			r.op = id;
			const auto e = estimates.find (id);
			if (e != estimates.end ())
				r.estimate = e->second;
			const operation *o = Find (id);
			if (o == nullptr)
				{
				++out.missing;
				Util::Log (part, L"regen: op " + std::to_wstring (id) + L" is not in the part any more - skipped");
				continue;
				}
			if (!o->db.nci_flag)
				{
				// Its toolpath is up to date already: only its time is read.
				++out.clean;
				Util::Log (part, L"regen: op " + std::to_wstring (id) + L" is regenerated already - only its cycle time is read");
				out.ops.push_back (r);
				continue;
				}
			ent opEnt;
			if (!opList.DatabaseRetrieve (id, opEnt))
				{
				++out.failed;
				r.failed = true;
				Util::Log (part, L"regen: op " + std::to_wstring (id) + L" FAILED - could not read it from the database");
				out.ops.push_back (r);
				continue;
				}
			DB_LIST_ENT_PTR eptr = nullptr;
			bool succf = false;
			const ULONGLONG t0 = GetTickCount64 ();
			operation_manager (&opEnt.u.op, OPMGR_NCI_REGEN, &eptr, &succf);
			const double took = static_cast<double> (GetTickCount64 () - t0) / 1000.0;
			Util::Log (part, L"regen: op " + std::to_wstring (id) + L" - operation_manager (OPMGR_NCI_REGEN) said "
							 + (succf ? L"done" : L"FAILED") + L", " + Seconds (took) + L", eptr "
							 + (eptr != nullptr ? L"set" : L"null"));
			if (succf)
				++out.regenerated;
			else
				{
				++out.failed;
				r.failed = true;
				}
			out.ops.push_back (r);
			}
		out.seconds = static_cast<double> (GetTickCount64 () - start) / 1000.0;

		// Still at it in the background: the toolpaths are not done, and their times
		// would be the old ones'.
		const long still = OperationInProcess (0, false);
		if (still != 0)
			{
			out.busy = true;
			Util::Log (part, L"regen: Mastercam is still regenerating in the background (OperationInProcess "
							 + std::to_wstring (still) + L") - cycle times not read");
			return out;
			}

		// Mastercam's own cycle time of each, as the dump reads it - only for one no
		// longer marked dirty, in the operation list or in the database (both said,
		// so a real run shows which one the regeneration brings up to date).
		const bool noTime = (Settings::Diag () & 1) != 0;
		for (History::Regenerated &r : out.ops)
			{
			operation *o = Find (r.op);
			if (o == nullptr || r.failed)
				continue;
			ent fresh;
			const bool inDb = opList.DatabaseRetrieve (r.op, fresh);
			if (o->db.nci_flag || (inDb && fresh.u.op.db.nci_flag))
				{
				++out.stillDirty;
				Util::Log (part, L"regen: op " + std::to_wstring (r.op) + L" is still marked dirty (list "
								 + (o->db.nci_flag ? L"dirty" : L"clean") + L", database "
								 + (!inDb ? L"not read" : fresh.u.op.db.nci_flag ? L"dirty" : L"clean")
								 + L") - its cycle time is not read");
				continue;
				}
			if (noTime)
				continue;
			const double t = CalcCycleTime (o, false);
			if (t > 0)
				r.mastercam = t;
			Util::Log (part, L"regen: op " + std::to_wstring (r.op) + L" - Mastercam's cycle time "
							 + (t > 0 ? History::Hms (t) : std::wstring (L"(none)")) + L", the estimate "
							 + (std::isnan (r.estimate) ? std::wstring (L"(none)") : History::Hms (r.estimate)));
			}
		Util::Log (part, L"regen: " + std::to_wstring (out.regenerated) + L" regenerated, " + std::to_wstring (out.clean)
						 + L" already, " + std::to_wstring (out.failed) + L" failed, " + std::to_wstring (out.missing)
						 + L" gone, " + std::to_wstring (out.stillDirty) + L" still dirty, " + Seconds (out.seconds)
						 + (noTime ? L" (diag 1: no cycle times)" : L""));
		return out;
		}

	void Keep (const std::filesystem::path &part, History::Book &history, const Outcome &o, double partEstimate)
		{
		if (o.ops.empty () || (o.regenerated == 0 && o.busy))
			return;
		if (!history.Add (History::RegenRecords (o.ops, partEstimate, History::Now ())))
			Util::Log (part, L"history: could not write " + history.file.wstring ());
		}

	std::wstring Said (const Outcome &o)
		{
		if (o.busy && o.regenerated == 0 && o.failed == 0)
			return L"Mastercam is regenerating operations already, so nothing was started. When it is done, run "
				   L"\"Lathe params - regenerate last load\".";
		std::wstring s = L"Regenerated " + std::to_wstring (o.regenerated) + L" operation(s) in "
						 + History::Hms (o.seconds) + L".";
		if (o.clean > 0)
			s += L" " + std::to_wstring (o.clean) + L" had been regenerated already.";
		if (o.failed > 0)
			s += L" " + std::to_wstring (o.failed) + L" FAILED - see ParamTable.log.";
		if (o.missing > 0)
			s += L" " + std::to_wstring (o.missing) + L" are no longer in the part.";
		if (o.busy)
			return s + L"\r\n\r\nMastercam is still regenerating in the background, so the cycle times were not read. "
					   L"When it is done, run \"Lathe params - regenerate last load\": it reads them without "
					   L"regenerating again.";
		double est = 0, mc = 0;
		int both = 0;
		for (const History::Regenerated &r : o.ops)
			if (!std::isnan (r.estimate) && !std::isnan (r.mastercam))
				{
				est += r.estimate;
				mc += r.mastercam;
				++both;
				}
		if (both > 0)
			{
			const double d = std::round (mc) - std::round (est);
			wchar_t pct[32] = L"";
			if (est > 0)
				swprintf_s (pct, L" (%+.1f%%)", std::round (d / est * 1000.0) / 10.0 + 0.0);
			s += L"\r\n\r\nMastercam's cycle time for them: " + History::Hms (mc) + L". The estimate was " + History::Hms (est)
				 + (d == 0 ? std::wstring (L" - the same.") : L" - Mastercam's is " + History::Hms (std::fabs (d))
															  + (d > 0 ? L" longer" : L" shorter") + pct + L".")
				 + L"\r\nKept in the part's history: the next dump shows it on the History sheet.";
			}
		if (o.stillDirty > 0)
			s += L"\r\n\r\n" + std::to_wstring (o.stillDirty) + L" still show as needing regeneration - check them in the "
				 L"Operation Manager.";
		return s;
		}

	int LastLoad ()
		{
		const std::filesystem::path part = Util::PartFile ();
		if (part.empty ())
			{
			Util::Say (L"Save the part first - there is no open part to regenerate operations in.");
			return 0;
			}
		const std::wstring name = part.filename ().wstring ();

		// ---- The last load: the part's history has it with the estimate its sheet
		// showed for each op; ParamTable.log has every load - a load from before the
		// history, or one it missed, is taken from there (without estimates).
		History::Book history = History::Open (part);
		const History::Record *load = History::Last (history.records, L"load");
		std::wstring log;
		const Undo::Last last = ReadFile (part.parent_path () / L"ParamTable.log", log) ? Undo::FindLast (log, name) : Undo::Last ();
		const bool logNewer = last.found && (load == nullptr || last.stamp.substr (0, 16) > load->when.substr (0, 16));
		std::vector<long> ops;
		std::map<long, double> estimates;
		double partEstimate = std::nan ("");
		std::wstring when;
		if (load != nullptr && !logNewer)
			{
			estimates = History::Estimates (*load);
			for (const auto &kv : estimates)
				ops.push_back (kv.first);
			partEstimate = History::Span (load->Get (L"cycle"), true).second;
			when = load->when;
			}
		if (ops.empty () && last.found)
			{
			std::set<long> seen;
			for (const Undo::Entry &e : last.entries)
				if (seen.insert (e.op).second)
					ops.push_back (e.op);
			when = last.stamp;
			}
		if (ops.empty ())
			{
			Util::Say (L"No load of " + name + L" was found - not in its history (" + history.file.filename ().wstring ()
					   + L") nor in ParamTable.log beside it - so there is nothing to regenerate.", MB_ICONINFORMATION);
			return 0;
			}
		// In the order they run.
		std::stable_sort (ops.begin (), ops.end (), [] (long a, long b) { return PlaceOf (a) < PlaceOf (b); });

		// ---- ASK, saying which and how long it can take. No is the default.
		std::wstring list;
		int dirty = 0, gone = 0, shown = 0;
		for (long id : ops)
			{
			const operation *o = Find (id);
			if (o == nullptr)
				++gone;
			else if (o->db.nci_flag)
				++dirty;
			if (++shown > 20)
				continue;
			list += L"\r\n    op " + std::to_wstring (id) + L"   "
					+ (o == nullptr ? std::wstring (L"(no longer in the part)")
									: std::wstring (o->comment, wcsnlen (o->comment, COMMENT_SIZE))
										  + (o->db.nci_flag ? L"   - needs regenerating" : L"   - regenerated already"));
			}
		if (shown > 20)
			list += L"\r\n    ... and " + std::to_wstring (shown - 20) + L" more";
		if (gone == static_cast<int> (ops.size ()))
			{
			Util::Say (L"The operations the last load of " + name + L" changed are no longer in the part:" + list,
					   MB_ICONINFORMATION);
			return 0;
			}
		const std::wstring head = L"The last load of " + name + (when.empty () ? L"" : L" (" + when + L")") + L" changed "
								  + std::to_wstring (ops.size ()) + (ops.size () == 1 ? L" operation:" : L" operations:") + list;
		const std::wstring after = L"\r\n\r\nAfterwards each operation's Mastercam cycle time is put beside the estimate "
								   L"its sheet showed, in the part's history (the History sheet at the next dump).";
		const std::wstring ask = dirty > 0
			? head + L"\r\n\r\nRegenerate the " + std::to_wstring (dirty) + (dirty == 1 ? L" that needs it now?" : L" that need it now?")
				  + L"\r\n\r\nThis can take a LONG time - twenty minutes or more on a big part - and Mastercam cannot be "
					L"used until it is done." + after
			: head + L"\r\n\r\nThey are all regenerated already. Read their Mastercam cycle times now?" + after;
		if (Util::Say (ask, MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2) != IDYES)
			return 0;

		Outcome o;
		{
		CWaitCursor wait;
		o = Run (part, ops, estimates);
		}
		Keep (part, history, o, partEstimate);
		Util::Say (Said (o) + L"\r\n\r\nSave the part to keep the regenerated toolpaths.",
				   o.failed > 0 || o.busy ? MB_ICONWARNING : MB_ICONINFORMATION);
		return 0;
		}
	}
