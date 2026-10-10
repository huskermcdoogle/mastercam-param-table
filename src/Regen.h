//
// Regen.h - regenerating the operations a load changed, ON PURPOSE. A
// regeneration can take a very long time - twenty minutes on one real part - so
// it is never done on the side: only when the load preview's box is ticked (it
// starts unticked every time), or by "Lathe params - regenerate last load",
// which asks first, with No as the default.
//
// THE MASTERCAM CALL is the C-Hook operation manager's own:
//
//   operation_manager (&op, OPMGR_NCI_REGEN, &eptr, &succf)
//
// (Mill\AssocHok_CH.h: "regenerate the nci section for the operation"; exported
// by MCMill.dll) - the one operation regeneration the 2026 SDK declares. It is
// made only for an operation that is dirty (db.nci_flag): a clean one has its
// toolpath already. Each call is logged with what Mastercam said back and how
// long it took, so a real run shows what it did. Nothing is started while
// Mastercam is regenerating something in the background (OperationInProcess);
// if it still is after the calls, the cycle times are not read - they would be
// the old toolpaths'.
//
// AFTER IT, each operation's own cycle time (CalcCycleTime - what the dump reads
// as cycle_time) goes into the part's history beside the estimate the load's
// sheet showed for it: the check on the estimate.
//
#pragma once

#include "History.h"

#include <filesystem>
#include <map>
#include <string>
#include <vector>

namespace Regen
	{
	/// What a regeneration did.
	struct Outcome
		{
		std::vector<History::Regenerated> ops;	//!< each op asked for that is in the part, in that order
		int regenerated = 0;		//!< regenerated now
		int clean = 0;				//!< regenerated already - only read
		int failed = 0;				//!< Mastercam said it failed
		int missing = 0;			//!< no longer in the part
		int stillDirty = 0;			//!< still marked dirty after it - not read
		double seconds = 0;			//!< how long the regenerations took
		bool busy = false;			//!< Mastercam was, or still is, regenerating in the background
		};

	/// Regenerate those of `ops` that need it, then read each one's Mastercam cycle
	/// time, with `estimates` (op -> the sheet's estimated seconds) beside it.
	/// Every step goes to ParamTable.log.
	Outcome Run (const std::filesystem::path &part, const std::vector<long> &ops,
				 const std::map<long, double> &estimates);

	/// The outcome in the part's history: the ops together (with the whole part,
	/// `partEstimate` - the part's estimate after the load), then each op.
	void Keep (const std::filesystem::path &part, History::Book &history, const Outcome &o, double partEstimate);

	/// What to tell the person about it, in a sentence or three.
	std::wstring Said (const Outcome &o);

	/// "Lathe params - regenerate last load": the operations the part's last load
	/// changed, listed, with a warning about the time - default No - then Run,
	/// Keep, and what happened.
	int LastLoad ();
	}
