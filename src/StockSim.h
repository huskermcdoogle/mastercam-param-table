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

namespace StockSim
	{
	/// Simulate every lathe machine group's operations in Operation Manager
	/// order and log one line per operation, plus the stock and tool shapes found.
	void Run (const std::filesystem::path &part);
	}
