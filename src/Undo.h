//
// Undo.h - putting the last load's old values back, from ParamTable.log. No
// Mastercam SDK.
//
// THE LOG IS THE RECORD. A load writes every old value to the log before it
// writes the operation, so the log is what an undo reads - there is no second
// copy to drift from it. The lines a load writes are made here, beside the
// code that reads them, for the same reason.
//
// One log serves every part in its folder, so a load opens with a line naming
// its part:
//
//   2026-10-09 10:11:12  load begin: shaft.mcam  <-  shaft_lathe_params_1.xlsx
//   2026-10-09 10:11:12  op 12 ROUGH  feed  0.3 -> 0.25
//   2026-10-09 10:11:12  op 12 ROUGH  feed_mode  per rev -> per min
//   2026-10-09 10:11:13  load: 1 operation(s) written, 0 failed
//
// A load from before this line existed cannot be undone: nothing in its lines
// says which part it was.
//
// THE RULES, each because the alternative is a silent wrong value:
//
//  * ONLY A VALUE STILL AS THE LOAD LEFT IT IS RESTORED. A value changed since
//    (by hand, or by another load) is somebody's later decision; putting the
//    old one over it would undo that, not the load.
//  * ONE CHANGED VALUE HOLDS BACK ITS WHOLE OPERATION. Half an operation
//    restored beside half of a later edit is a state nobody chose - a feed
//    from before the load with a feed mode from after it.
//  * A VALUE ALREADY BACK TO ITS OLD SETTING IS LEFT ALONE, so an undo run
//    twice does nothing the second time.
//
#pragma once

#include "Plan.h"

#include <string>
#include <vector>

namespace Undo
	{
	// ---- The lines a load writes ------------------------------------------

	/// The line that opens a load: which part, from which sheets.
	std::wstring BeginLine (const std::wstring &partName, const std::wstring &files);

	/// One value about to be written: "op 12 ROUGH  feed  0.3 -> 0.25".
	std::wstring ChangeLine (long op, const std::wstring &type, const std::wstring &column,
							 const std::wstring &from, const std::wstring &to);

	/// A value on one log line: a manual entry's line breaks shown as the
	/// preview shows them, so one change stays one line.
	std::wstring OneLine (const std::wstring &v);

	/// OneLine undone: the shown line breaks back to real ones, CR LF unless
	/// `like` (the value as it is now) uses bare LF.
	std::wstring FromOneLine (const std::wstring &v, const std::wstring &like);

	// ---- Reading the log back ---------------------------------------------

	/// One value a load wrote.
	struct Entry
		{
		long op = 0;
		std::wstring type;			//!< "ROUGH"
		std::wstring column;		//!< "feed"
		std::wstring both;			//!< "0.3 -> 0.25" as logged: old, then new
		};

	/// The most recent load of one part.
	struct Last
		{
		bool found = false;
		std::wstring stamp;			//!< when it began, "2026-10-09 10:11:12"
		std::wstring files;			//!< the sheets it loaded
		std::wstring end;			//!< its closing line ("load: 3 operation(s) written ..."), "" if none
		std::vector<Entry> entries;	//!< one per operation and column, the last written

		/// A log that has loads without a begin line (written before undo
		/// existed) - said when nothing is found, so the refusal names why.
		bool olderLoads = false;
		};

	/// The most recent load of `partName` (a file name, matched ignoring case)
	/// in the log's text.
	Last FindLast (const std::wstring &log, const std::wstring &partName);

	/// The two halves of an entry's "old -> new". A value of free text could
	/// itself hold " -> ", so every split is offered, the first one first.
	std::vector<std::pair<std::wstring, std::wstring>> Splits (const std::wstring &both);

	// ---- Deciding what to restore -----------------------------------------

	/// One kind of operation as it stands now.
	struct Kind
		{
		const Plan::Schema *schema = nullptr;
		std::vector<Plan::Current> now;
		};

	struct Restore
		{
		std::wstring type;
		Plan::Change change;		//!< from = the value now, to = the value before the load
		};

	struct Skip
		{
		long op = 0;
		std::wstring type;
		std::wstring column;		//!< "" = the whole operation
		std::wstring why;
		};

	struct Result
		{
		std::vector<Restore> restores;
		std::vector<Skip> skipped;	//!< changed since, or gone - left alone
		std::vector<Entry> already;	//!< back to the old value already
		};

	Result Make (const Last &last, const std::vector<Kind> &kinds);
	}
