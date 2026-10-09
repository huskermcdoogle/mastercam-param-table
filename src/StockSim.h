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
		bool hasAir = false;	//!< airPct is known (simulated)
		bool fromBoundary = false;	//!< removed is Mastercam's stock boundary difference
		double simRemoved = -1;		//!< the simulation's own figure (-1: not simulated)
		double mcRemoved = -1;		//!< Mastercam's: its stock boundary before less after (-1: none)
		bool simOk = false;			//!< the simulation covered the op
		};

	/// Simulate every lathe machine group's operations in Operation Manager
	/// order: what each removes, by op_idn. `log` writes the probe's lines too.
	/// The shape the simulation sweeps for a lathe tool, measured from its outline
	/// (the tool's cut boundary): round with a diameter, or a polygon with the
	/// included angle at its nose. For checking against the insert's own data.
	struct ToolShape
		{
		bool ok = false;			//!< there is an outline (not just a nose circle)
		bool round = false;
		double size = 0;			//!< round: diameter; polygon: the larger extent
		double noseAngle = 0;		//!< polygon: degrees at the nose
		double noseRadius = 0;
		int lines = 0, arcs = 0;	//!< the cut boundary: how many lines and arcs
		double maxArcR = 0, maxArcSweep = 0;	//!< its largest arc: radius, sweep (degrees)
		};
	ToolShape ShapeOfTool (long slot);

	/// simulate = false: no sweeping - each op's removed volume is the difference
	/// of Mastercam's own stock boundaries before and after it (quick; no air share).
	std::map<long, Result> Run (const std::filesystem::path &part, bool log, bool simulate);
	}
