#pragma once

#include "Settings.h"
#include "Xlsx.h"

#include <filesystem>
#include <string>

namespace Dump
	{
	/// "Lathe params - dump to Excel": the open part's operations to a workbook -
	/// which ones, where and what it is called asked in the dump window, the result
	/// said, and Excel opened on it.
	int Run ();

	/// What one dump wrote - for the batch dump, which says it per part and puts
	/// the parts side by side in one shop workbook (Shop.h).
	struct Outcome
		{
		std::filesystem::path file;		//!< the workbook (also when it could not be written)
		size_t ops = 0;					//!< operations dumped
		std::wstring why;				//!< why not, in words
		std::wstring skipped;			//!< kinds of op not read yet ("THREAD x2, POINT x1")
		std::wstring material;			//!< the machine group's stock material, "" = none set
		Xlsx::Sheet sheet;				//!< the workbook as written, cached values and all
		};

	/// The open part dumped WHOLE - every op - with these settings, exactly as Run
	/// dumps it (the same function writes both), asking nothing, saying nothing and
	/// opening nothing: what "dump a folder of parts" does to each part. False (and
	/// `out.why`) when there was nothing to dump or the workbook could not be written.
	bool Quiet (const Settings::Dump &settings, Outcome &out);
	}
