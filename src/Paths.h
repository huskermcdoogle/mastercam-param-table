//
// Paths.h - walking an operation's NCI: how far the tool moves, and where.
//
// The groundwork for estimating cycle time from the sheet's own feeds and
// speeds. Under CSS the spindle speed changes with diameter, so a length is
// not enough - where along X it was cut matters - and the walk keeps that.
//
// LATHE NCI: X is a RADIUS (confirmed on a real part: 7.1 on a 14" OD).
//
#pragma once

#include "Estimate.h"

#include <string>
#include <vector>

struct operation;

namespace Paths
	{
	struct Totals
		{
		bool ok = false;			//!< the operation had NCI to walk
		long moves = 0;				//!< feed + rapid moves
		long sections = 0;			//!< 1002 lines - each sets the spindle speed for what follows
		long holes = 0;				//!< canned drill holes
		double drillLength = 0;		//!< their feed travel (feed plane to bottom), in cutLength too
		double cutLength = 0;		//!< feed moves (lines and arcs)
		double rapidLength = 0;
		double arcLength = 0;		//!< the part of cutLength that is arcs
		long arcs = 0;
		long arcsOverHalf = 0;		//!< arcs sweeping more than 180 degrees - rare in
									//!< turning; many would mean the direction is read backwards
		long planes[3] = { 0, 0, 0 };	//!< the arcs' plane codes as stored, for the record
		double min[3] = { 0, 0, 0 }, max[3] = { 0, 0, 0 };	//!< over every point visited

		/// Feed time at the NCI's own feeds and speeds: each section's speed (CSS
		/// turned to RPM at each move's diameter, capped), each move's feed.
		/// Seconds; rapids and dwells excluded.
		double feedSeconds = 0;

		/// The feed moves by feed value and spindle setting, banded by diameter:
		/// what the sheet's live estimate is built from. A dynamic mill's back
		/// feed is its own group.
		std::vector<Estimate::Group> groups;
		};

	/// Walk one operation's NCI. Totals::ok is false when it has none (needs
	/// regenerating, or a kind with no moves).
	Totals Walk (operation &op);

	/// Every NCI line of the operation as CSV rows (gcode, type, plane, end point,
	/// centre, feed ...) - raw, for working out how a move should be read.
	std::wstring Listing (operation &op);

	/// One line for the log: lengths, ranges, and the estimate beside Mastercam's.
	std::wstring Describe (const operation &op, const Totals &t, double mastercamSeconds);
	}
