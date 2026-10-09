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
// INSERT NAMES. The dump labels each tool's insert itself (its geometry, a
// PrimeTurning or groove type ...). The file records that label as `detected`;
// only a name typed over it is kept as the tool's `insert` - so better labels in
// later versions still come through where nobody typed one.
//
// The file is plain text, UTF-8, one "key = value" per line under [sections]:
//   [batch]          qty
//   [insert <name>]  edges, cost, parts_per_edge
//   [tool <number>]  edge_life, insert, detected
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
		};
	struct Tool
		{
		std::wstring edgeLife;		//!< m:ss as typed
		std::wstring insert;		//!< a name typed over the detected one
		std::wstring detected;		//!< what the last dump labelled it
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
	/// `c`: a value filled in there replaces the file's; a blank one leaves it.
	/// False (and `c` untouched) when the workbook cannot be read.
	bool Harvest (const std::filesystem::path &workbook, Config &c, std::wstring &why);

	/// The newest dump workbook of this part beside it or in `folder` (file names
	/// starting with the part's name and holding "_params"), "" when none.
	std::filesystem::path NewestDump (const std::filesystem::path &part, const std::filesystem::path &folder);
	}
