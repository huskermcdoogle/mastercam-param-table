//
// Estimate.h - feed time as a live Excel formula of the sheet's own cells.
// No Mastercam SDK.
//
// THE MODEL (it reproduces Mastercam's cycle time to within seconds on a real
// part): a feed move of length L at radius r takes
//
//     per minute:  60 L / F
//     per rev:     60 L / (F * RPM(r)),  RPM(r) = CSS ? min(cap, k S / (pi D)) : S
//
// with D = 2r and k = 12 (SFM, inches) or 1000 (m/min, mm). An operation's moves
// are collected into GROUPS - one per feed value and spindle setting - and each
// group's moves into a few diameter BANDS (total length, length-weighted mean
// diameter), which is all the formula needs. A group's feed and speed are tied
// to the sheet cells they came from, so editing a cell moves the estimate.
//
#pragma once

#include <string>
#include <utility>
#include <vector>

namespace Estimate
	{
	struct Band
		{
		double len = 0;			//!< total length of the band's moves
		double lenDia = 0;		//!< sum of length x diameter
		};

	/// One feed value under one spindle setting, as the NCI ran it.
	struct Group
		{
		double feed = 0;		//!< as the NCI holds it: negative per rev, positive per minute
		double speed = 0;		//!< RPM, or surface speed when css
		bool css = false;
		double cap = 0;			//!< max RPM, 0 = none
		bool mm = false;
		/// Lathe moves by direction: 1 mostly along Z (axial), 2 mostly along X
		/// (radial), 0 not split. Prime turning feeds the two differently.
		int axis = 0;
		std::vector<Band> bands;
		};

	/// Where a group's settings live on the sheet. An empty reference means
	/// "not on the sheet" - the dumped value is written into the formula.
	struct Cells
		{
		std::wstring feed;		//!< the feed's size (always positive)
		std::wstring feedMode;	//!< a cell holding "per rev" / "per min"
		std::wstring speed;
		std::wstring css;		//!< a cell holding "CSS"/"RPM", or 1/0
		bool cssIsWord = true;	//!< css cell holds "CSS"/"RPM" (else 1/0)
		std::wstring cap;
		};

	/// (length, radius) pieces to at most `maxBands` bands of roughly equal length,
	/// banded by diameter.
	std::vector<Band> Bands (std::vector<std::pair<double, double>> pieces, size_t maxBands);

	/// Seconds for a group at its own (dumped) values - what the formula gives
	/// before anything is edited.
	double Seconds (const Group &g);

	/// An Excel expression for the group's seconds, reading the cells given.
	std::wstring Term (const Group &g, const Cells &c);
	}
