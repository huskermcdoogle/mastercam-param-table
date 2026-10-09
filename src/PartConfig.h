//
// PartConfig.h - what a person types into a dump that Mastercam does not hold -
// per insert its edges, cost and parts per edge; per tool its edge life and a
// corrected insert name; the batch quantity - kept in a small file beside the
// part ("<part>.ptconfig"), so a new dump starts from it and nothing typed is
// lost, whether or not the edited workbook was ever loaded back.
//
// THE FLOW. At a dump, before writing: the newest earlier workbook of this part
// (if it is newer than the file) is read - its Tools and Summary sheets - and
// what was typed there goes into the file; then the file's values fill the new
// workbook. A load does the same from the workbook it loads. No macros: the
// workbook only has to be SAVED.
//
// WHAT THE DUMP FILLS IN ITSELF - an insert's label, its edges, a tool's edge
// life - is recorded as `detected` / `edges_detected` / `life_detected`; only a
// value typed OVER it is kept. So better guesses in later versions, and a changed
// insp_time, still come through where nobody typed anything. A value with no such
// record (a workbook from before) is not taken: it cannot be told from a guess.
// A dump of the WHOLE part drops tools and inserts it no longer uses.
//
// The file is plain text, UTF-8, one "key = value" per line under [sections]:
//   [batch]          qty
//   [insert <name>]  edges, cost, parts_per_edge, edges_detected
//   [tool <number>]  edge_life, insert, detected, life_detected
//
#pragma once

#include <filesystem>
#include <map>
#include <string>

namespace PartConfig
	{
	struct Insert
		{
		std::wstring edges, cost, partsPerEdge;		//!< as typed ("" = not set)
		std::wstring edgesDetected;					//!< the edges the last dump filled in
		};
	struct Tool
		{
		std::wstring edgeLife;		//!< m:ss as typed
		std::wstring insert;		//!< a name typed over the detected one
		std::wstring detected;		//!< what the last dump labelled it
		std::wstring lifeDetected;	//!< the edge life the last dump filled in
		};
	struct Config
		{
		std::wstring batchQty;
		std::map<std::wstring, Insert> inserts;		//!< by insert name
		std::map<std::wstring, Tool> tools;			//!< by tool number
		};

	/// "<folder>\<stem>.ptconfig" for a part file.
	std::filesystem::path PathFor (const std::filesystem::path &part);

	/// Read the file (an empty config when there is none or it cannot be read).
	Config Load (const std::filesystem::path &file);
	/// Write it. False when it could not be written.
	bool Save (const std::filesystem::path &file, const Config &c);

	/// What was typed into a dump workbook's Tools and Summary sheets, merged into
	/// `c`: cost and parts per edge as they stand (a dump writes the kept ones
	/// back, so blank = cleared); edges, edge life and insert names only where
	/// they differ from what that dump filled in.
	/// False (and `c` untouched) when the workbook cannot be read.
	bool Harvest (const std::filesystem::path &workbook, Config &c, std::wstring &why);

	/// The newest dump workbook of this part beside it or in `folder` (file names
	/// starting with the part's name and holding "_params"), "" when none.
	std::filesystem::path NewestDump (const std::filesystem::path &part, const std::filesystem::path &folder);
	}
