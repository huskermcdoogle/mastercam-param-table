//
// Shop.h - the batch dump's SHOP WORKBOOK: every part of a folder dumped, then put
// side by side - where continuous improvement across the shop starts, rather than
// one part at a time. No Mastercam SDK: each part is read from its dump AS
// WRITTEN (the sheet's cached values, found by column name and by heading), so all
// of it is tested at a command prompt (tests\shop_test.cpp).
//
// THE SHEETS
//   Parts    one row per part: dumped or why not, its ops, cycle time, cutting
//            time, insert flips and insert cost per part (when its .ptconfig
//            holds the costs - the dump fills them in from there).
//   Inserts  one row per insert, per kind of op, per material: the parts and
//            tools that use it, and the range and median of its surface speed,
//            feed per rev, depth of cut and edge time - what a shop STANDARDS
//            LIBRARY is built from. Its column names are fixed (another tool
//            reads them): insert, op_kind, material, sfm_median, feed_median,
//            depth_median, edge_time_median, n_ops, parts.
//   Ops      every op of every part: speed and feed as dumped and in inch units,
//            depth of cut, times, flips, edge time, MRR, air.
//   About    what each column holds, and its units.
//
// ONE SET OF UNITS. Parts in inches and in millimetres are compared side by side,
// so the speeds, feeds and depths the Inserts sheet works from are in inch
// units: surface feet per minute, inches per rev, inches (a metric part's m/min,
// mm/rev and mm converted). The Ops sheet keeps each op's own values too.
//
// EDGE TIME is a tool's, as the Program check works it out: its cutting time per
// part over its insert flips per part (or the longest it ran between flips, when
// that is longer); with parts per edge typed instead, cut time x parts per edge;
// unknown when its edge outlasts the part and nobody said by how much.
//
#pragma once

#include "Xlsx.h"

#include <ctime>
#include <filesystem>
#include <limits>
#include <string>
#include <vector>

namespace Shop
	{
	const double kNone = std::numeric_limits<double>::quiet_NaN ();

	/// One operation of one part, as the shop workbook lists it.
	struct Op
		{
		std::wstring part;				//!< the part's name (its file name, no extension)
		std::wstring op, type, comment, tool, insert, units, needsRegen;
		/// As dumped, in the part's own units: the speed (SFM / m/min when CSS, else
		/// RPM) and its mode, the feed and its mode, the depth of cut.
		std::wstring speed, speedMode, feed, feedMode, depth;
		/// In inch units, worked out: surface speed (ft/min) at the op's mean cutting
		/// diameter, feed per rev (in), depth of cut (in). kNone = cannot be said.
		double sfm = kNone, feedIpr = kNone, depthIn = kNone;
		double estSeconds = kNone, cutSeconds = kNone, flips = kNone;
		double edgeSeconds = kNone;		//!< its tool's edge time (see the top)
		double mrr = kNone, airPct = kNone;
		std::wstring mrrUnits;			//!< "in3/min" or "cm3/min"
		};

	/// One part: dumped, or why not.
	struct Part
		{
		std::wstring name;				//!< the part file's name, no extension
		std::wstring file;				//!< its path from the folder dumped ("Lathe\\Shaft.mcam")
		bool ok = false;
		std::wstring why;				//!< why it was not dumped; or a note on one that was
		std::wstring workbook;			//!< its own dump's file name
		std::wstring material;			//!< the stock material, "" = unknown
		std::wstring units;				//!< "in" / "mm" (its ops'), "" = none said
		std::wstring skipped;			//!< kinds of op not read yet ("THREAD x2")
		std::vector<Op> ops;
		double cycleSeconds = kNone, cutSeconds = kNone, flips = kNone;
		size_t regen = 0;				//!< ops that needed regenerating when dumped
		size_t inserts = 0, costed = 0;	//!< inserts it uses, and how many have a cost
		double insertCost = kNone;		//!< insert cost per part over the costed inserts
		};

	/// A part's figures, read from its dump as written - the main sheet's rows (by
	/// column name), its tools' inserts (Tools page, by heading) and its inserts
	/// table. `p.name` and `p.material` are the caller's; the rest is filled in.
	void Read (const Xlsx::Sheet &s, Part &p);

	/// How long one edge of a tool cuts, in seconds (see the top): `cut` and `flips`
	/// per part, `longest` between flips (0 = none), `partsPerEdge` typed (0 = none).
	/// kNone when it cannot be said.
	double EdgeSeconds (double cut, double flips, double longest, double partsPerEdge);

	/// The smallest, middle and largest of some values (NaN ones left out); n = how
	/// many were counted. The middle of an even count is the mean of the two.
	struct Spread
		{
		size_t n = 0;
		double min = kNone, median = kNone, max = kNone;
		};
	Spread SpreadOf (std::vector<double> values);

	/// One row of the Inserts sheet: an insert at one kind of op, on one material.
	struct InsertRow
		{
		std::wstring insert, opKind, material;
		size_t nOps = 0;
		std::vector<std::wstring> parts;	//!< in first-seen order
		std::wstring tools;				//!< "Shaft: T3, T5; Flange: T1"
		Spread sfm, feed, depth, edge;	//!< inch units; edge time in seconds, once per tool
		double flips = 0;				//!< insert flips per part, added over its ops
		};

	/// The Inserts sheet's rows: the dumped parts' ops that have an insert, grouped
	/// by insert (case aside), kind of op and material - in that order.
	std::vector<InsertRow> ByInsert (const std::vector<Part> &parts);

	/// Every part file in a folder (and its subfolders, if asked), in name order.
	/// `ext` is the part extension (".mcam"); `limit` stops the search at that many
	/// (0 = no limit) and `more` says whether there were more. A folder that cannot
	/// be read is passed over.
	std::vector<std::filesystem::path> FindParts (const std::filesystem::path &folder, bool subfolders,
												  const std::wstring &ext, size_t limit, bool &more);

	/// The shop workbook's file name for a day: "Shop_compare_<yyyymmdd>.xlsx".
	/// (Not "..._params": a part whose name starts the same would take it for one
	/// of its own dumps - PartConfig::NewestDump.)
	std::wstring FileName (std::time_t when);

	/// The workbook's sheets: Parts, Inserts, Ops, About. `title` heads About
	/// ("Shop compare - <folder> - <date>").
	std::vector<Xlsx::Table> Tables (const std::vector<Part> &parts, const std::wstring &title);
	}
