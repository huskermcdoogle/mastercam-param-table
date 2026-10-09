//
// StockSim.h - PROBE: the material each lathe operation really removes, from a
// 2D stock simulation of the toolpath (not Mastercam's stock models).
//
// The stock starts as the machine group's lathe stock (cylinder OD/ID/length).
// Each operation's feed moves sweep the tool's insert outline through a raster
// of the XZ half-plane; every cell newly cleared is a ring of 2*pi*r*cellArea
// (Pappus). Output goes to ParamTable.log only (diag bit 4) - nothing on the sheet.
//
#pragma once

#include <filesystem>
#include <map>

namespace StockSim
	{
	/// One operation's result.
	struct Result
		{
		double removed = 0;		//!< in^3 (cm^3 for a metric part) removed by the op
		double airPct = 0;		//!< % of its feed time cutting nothing
		bool ok = false;		//!< false: drill cycles, mill moves, or after a stock flip -
								//!< what it removes is not (all) simulated
		};

	/// Simulate every lathe machine group's operations in Operation Manager
	/// order: what each removes, by op_idn. `log` writes the probe's lines too.
	std::map<long, Result> Run (const std::filesystem::path &part, bool log);
	}
