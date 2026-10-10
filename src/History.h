//
// History.h - a part's improvement history, kept beside it in "<part>.pthistory":
// every dump, load, undo and regeneration, what was measured on the machine,
// and the reasons typed for the changes - so whether a program change saved
// time and inserts is on record, not remembered. The dump shows it as the
// workbook's History sheet. No Mastercam SDK.
//
// THE FILE is plain UTF-8 text (with a BOM, so Notepad shows any character
// right), one record per line, oldest first. The tool only ever ADDS lines; a
// line starting with "#" is a comment, and a line it cannot read is skipped.
// Delete the file to start a new history. A record is fields split by " | ":
// when, what kind, then "name: value" fields:
//
//   2026-10-10 09:14 | dump | file: shaft_params_1.xlsx | ops: 42 | whole part: yes | cycle: 26:32:07 | flips: 168 | insert cost: 12.25 | batch: 40
//   2026-10-10 11:02 | load | file: shaft_params_1.xlsx | changes: 14 | ops: 6 | cycle: 26:32:07 -> 25:10:40 | flips: 168 -> 152 | insert cost: 12.25 -> 11.1 | batch: 40 | estimates: 12 0:04:12, 15 0:01:03
//   2026-10-10 11:02 | why | op: 12 | column: feed | change: 0.3 -> 0.25 | why: chatter on the shoulder
//   2026-10-11 08:30 | undo | load: 2026-10-10 11:02:13 | restored: 14 | complete: yes
//   2026-10-12 14:20 | regen | ops: 2 | estimate: 0:05:15 | mastercam: 0:05:21 | part estimate: 25:10:40 | part mastercam: 25:10:46
//   2026-10-12 14:20 | regen | op: 12 | estimate: 0:04:12 | mastercam: 0:04:15
//   2026-10-13 07:05 | actual | measured: 2026-10-12 | cycle: 25:40:00 | parts: 40 | inserts: 31 | estimate: 25:10:40 | note: first batch after the change
//
//   dump    a workbook written: its operations, and the part's figures in it -
//           cycle time (the sheet's estimate, which starts at Mastercam's own),
//           insert flips and insert cost per part, the batch quantity.
//   load    a workbook loaded: the changes written, and the figures per part
//           BEFORE -> AFTER. Before is the workbook's Dumped sheet (as dumped);
//           after is what the sheet showed when it was loaded, for the
//           operations written (one whose changes were all left out counts as
//           dumped). `estimates` is each written op's estimated time after -
//           what a regeneration is measured against.
//   why     a reason typed in the Change report's Why column, kept once.
//           Written straight after its load, it belongs to that load.
//   undo    "undo last load" ran: how many values it put back, and whether that
//           was the whole load (complete: yes) - then the load's saving is gone.
//   regen   operations regenerated: per op, the load's estimate and Mastercam's
//           own cycle time after; and the ops together, with the whole part
//           (the part's estimate after the load, and the same with those ops at
//           Mastercam's figures). This is how the estimate is checked.
//   actual  what the part took on the machine, typed into the Summary's "From
//           the machine" cells: the date measured (`date typed: no` when the day
//           it was read stands in), cycle time per part, parts run, inserts used,
//           a note, and the sheet's estimate beside it then. Kept once: the same
//           date and figures again (the next dump shows the newest one again)
//           are the same measurement.
//
// Times are h:mm:ss; a cost has no currency (the workbook's is the computer's).
// A "|" typed into a text is written "¦" and a line break " ↵ ", so a record
// is always one line.
//
#pragma once

#include "Csv.h"
#include "Xlsx.h"

#include <cmath>
#include <filesystem>
#include <map>
#include <set>
#include <string>
#include <utility>
#include <vector>

namespace History
	{
	/// One line of the file.
	struct Record
		{
		std::wstring when;			//!< "2026-10-10 11:02"
		std::wstring kind;			//!< dump, load, why, undo, regen, actual - others are kept as read
		std::vector<std::pair<std::wstring, std::wstring>> fields;	//!< in the order written

		/// A field's value, "" when it has none.
		std::wstring Get (const std::wstring &name) const;
		/// Whether the record has the field at all.
		bool Has (const std::wstring &name) const;
		/// Set a field (replacing one of that name). An empty value is not written.
		Record &Set (const std::wstring &name, const std::wstring &value);
		};

	/// "<folder>\<stem>.pthistory" for a part file.
	std::filesystem::path PathFor (const std::filesystem::path &part);

	/// A record as its line in the file (no line end).
	std::wstring Line (const Record &r);

	/// A line of the file back to a record. False for a comment, a blank line, or
	/// a line with no kind.
	bool Parse (const std::wstring &line, Record &r);

	/// Every record of a file, oldest first (none when there is no file).
	std::vector<Record> Read (const std::filesystem::path &file);

	/// Add records at the end of the file - a new file starts with a few lines
	/// saying what it is. False when it could not be written.
	bool Append (const std::filesystem::path &file, const std::vector<Record> &records);

	/// The part's history, read once and added to as it goes.
	struct Book
		{
		std::filesystem::path file;
		std::vector<Record> records;

		/// Write `more` at the end of the file and keep them here. False (nothing
		/// kept) when the file could not be written.
		bool Add (const std::vector<Record> &more);
		};

	/// The history of a part: its file, read.
	Book Open (const std::filesystem::path &part);

	// ---- Times, dates, figures ------------------------------------------------

	/// A time as a person types it, in seconds: "4:30" (m:ss), "1:02:30" (h:mm:ss),
	/// "270" (seconds) - or a number under 1, a time Excel turned into a fraction of
	/// a day. NaN when it is none of those.
	double Seconds (const std::wstring &text);

	/// Seconds as h:mm:ss ("26:32:07", "0:04:12"); "" for NaN or less than 0.
	std::wstring Hms (double seconds);

	/// A change in seconds with its sign: "+0:00:32", "-1:21:27".
	std::wstring Change (double seconds);

	/// Now, as a record's `when`: "2026-10-10 11:02" (local time).
	std::wstring Now ();

	/// An Excel date (days since 1899-12-30) as "2026-10-12"; "" when it is not a
	/// date this side of 1900-03-01.
	std::wstring DateOfSerial (double serial);

	/// "2026-10-12" (or 2026/10/12) as an Excel date; NaN when it is not one.
	double SerialOfDate (const std::wstring &date);

	/// A figure for a record: rounded to the hundredth, no trailing zeros.
	std::wstring Figure (double v);

	/// "a -> b" split into its two numbers or times (`time`: h:mm:ss); NaN where a
	/// side is missing. A single value is both.
	std::pair<double, double> Span (const std::wstring &text, bool time);

	// ---- What a workbook says -------------------------------------------------

	/// The "From the machine" cells of a saved workbook's Summary.
	struct Actual
		{
		std::wstring cycle;			//!< h:mm:ss
		std::wstring measured;		//!< yyyy-mm-dd as typed (or as typed, if not a date); "" = not typed
		std::wstring parts, inserts, note;
		std::wstring estimate;		//!< h:mm:ss - the sheet's estimate beside it
		};

	/// Read them. False when no actual cycle time is typed (or it is not a time).
	bool ReadActual (const Xlsx::Grid &summary, Actual &a);

	/// The measurement typed into a workbook's Summary, as a record - none when
	/// nothing is typed, or when `have` holds it already (the same date and
	/// figures). A measurement with no date typed takes the day of `now`.
	std::vector<Record> Measured (const std::filesystem::path &workbook, const std::vector<Record> &have,
								  const std::wstring &now);

	/// The reasons typed in a workbook's Change report (its Why column), one record
	/// each - those `have` does not hold already. None without a Change report.
	std::vector<Record> Whys (const std::filesystem::path &workbook, const std::vector<Record> &have,
							  const std::wstring &now);

	/// The newest measurement in the history, or nullptr.
	const Record *NewestActual (const std::vector<Record> &records);

	/// The latest record of a kind, or nullptr.
	const Record *Last (const std::vector<Record> &records, const std::wstring &kind);

	/// A load record's `estimates`: op -> estimated seconds (NaN where it had none).
	std::map<long, double> Estimates (const Record &load);

	/// INSERT COST PER PART, from a dump workbook's sheets, worked out the way its
	/// Tools page does: each insert's flips (the flips_part of its tools' ops, a
	/// transform's copies counted) over its edges, times its cost - or 1 / parts
	/// per edge for an insert the ops flip less than once a part. `main` and
	/// `dumped` are the main and Dumped sheets' rows ([0] the column names); an op
	/// in `now` counts its main-sheet flips, any other its Dumped-sheet flips (its
	/// main-sheet ones when there is no Dumped sheet). NaN when no insert in the
	/// Tools page's inserts table has a cost typed in.
	double InsertCost (const std::vector<Csv::Row> &main, const std::vector<Csv::Row> &dumped,
					   const Xlsx::Grid &tools, const std::set<long> &now);

	/// The Tools page of a sheet about to be written, as the grid reading it back
	/// would give: the heading row 1, a row per tool, then the rows under them.
	Xlsx::Grid ToolsGrid (const Xlsx::Sheet &s);

	// ---- The records the tool writes ------------------------------------------

	/// A dump of `s` (as it will be written) to `file`.
	Record DumpRecord (const Xlsx::Sheet &s, const std::wstring &file, bool wholePart,
					   const std::wstring &batch, const std::wstring &when);

	/// What a load did. NaN = the sheet does not say.
	struct LoadFigures
		{
		std::wstring files;
		int changes = 0, ops = 0, failed = 0;
		double cycleWas = std::nan (""), cycleNow = std::nan ("");
		double flipsWas = std::nan (""), flipsNow = std::nan ("");
		double costWas = std::nan (""), costNow = std::nan ("");
		std::wstring batch;
		std::vector<std::pair<long, double>> estimates;	//!< each op written, its estimated seconds after
		};
	Record LoadRecord (const LoadFigures &f, const std::wstring &when);

	/// One regenerated operation: the load's estimate for it and Mastercam's own
	/// cycle time after the regeneration (NaN = not known / not read).
	struct Regenerated
		{
		long op = 0;
		double estimate = std::nan (""), mastercam = std::nan ("");
		bool failed = false;		//!< Mastercam said the regeneration failed
		};

	/// A regeneration's records: the ops together first (with the whole part:
	/// `partEstimate`, the part's estimate after the load, and the same with these
	/// ops at Mastercam's figures), then one per op - so on the History sheet the
	/// ops are listed under the line for them all.
	std::vector<Record> RegenRecords (const std::vector<Regenerated> &ops, double partEstimate,
									  const std::wstring &when);

	// ---- The History sheet ----------------------------------------------------

	/// The workbook's History page: what the loads saved so far (per part, per
	/// batch of `batchNow`), the last measurement and regeneration, then every
	/// record newest first in plain words - with the time saved per part and per
	/// batch at each load and a running total, and the estimate against what was
	/// measured. Printed landscape, one page wide.
	Xlsx::Sheet::Page Page (const std::vector<Record> &records, const std::wstring &partName,
							const std::wstring &historyFile, const std::wstring &batchNow);
	}
