//
// ToolPictures.h - a tool's picture as PNG, for the sheet's Tools page, and
// what the tool manager says about its insert and holder.
//
// Mastercam draws it as a bitmap - a lathe tool through the call its setup
// sheets use, a mill tool through the tool manager's bitmap factory - and it
// is converted to PNG here, since that is what a workbook embeds well.
//
// The ISO / ANSI code readers at the end use no SDK (inline, so the unit tests
// build them without Mastercam).
//
#pragma once

#include <cmath>
#include <cwctype>
#include <string>

namespace ToolPictures
	{
	/// The lathe tool in this tool-list slot, drawn, as PNG bytes and its size in
	/// pixels. False, with a reason, when there is no such tool or no picture.
	bool LatheTool (long slot, std::string &png, int &width, int &height, std::wstring &why);

	/// The same for the MILL tool in this mill tool-list slot. Never throws and
	/// never asks anything: a failure is false with a reason.
	bool MillTool (long slot, std::string &png, int &width, int &height, std::wstring &why);

	/// The insert name of the lathe tool in this slot ("" when none).
	std::wstring LatheInsert (long slot);

	/// The lathe tool's manufacturer code (a 3D tool's order code; "" when blank).
	std::wstring LatheMfgCode (long slot);

	/// The lathe tool's main insert as the tool manager defines it.
	struct InsertInfo
		{
		bool ok = false;			//!< the tool has an insert definition
		wchar_t shape = 0;			//!< ANSI shape code: C, D, V, T, S, R, W ... (0 = none / custom)
		double ic = 0, radius = 0, thickness = 0, width = 0, length = 0;	//!< in the part's units
		std::wstring grade;			//!< the insert grade's name
		bool custom = false;
		};
	InsertInfo LatheInsertInfo (long slot);

	/// The lathe tool's main holder as the tool manager defines it.
	///
	/// WHAT THE LIBRARIES SHOWED (Mastercam 2026's own, Sandvik's and Kennametal's
	/// lathe .tooldb files, read before this was written): `style` is the holder's
	/// ISO / ANSI style letter - PCLNR -> L, PDJNR -> J, PSSNR -> S, S25T-SDUCR -> U -
	/// with Mastercam's own button holders lower case (a, b). Mastercam's library
	/// DCGNR-164D is drawn as style L although its code says G. The side and end
	/// cutting edge angles are the drawing's and do NOT give the entering angle
	/// (an L-style holder stores end angles of 0, 10 or 52) - logged only. A
	/// PrimeTurning tool carries a stand-in holder (DCGNR / MWLNR): no angle there.
	struct HolderInfo
		{
		bool ok = false;			//!< the tool has a holder definition
		std::wstring name;			//!< the holder's name - often its code (MCLNR-164D)
		wchar_t style = 0;			//!< its style letter (ISO 5608 position 3)
		int type = -1;				//!< 0 general turning, 1 threading, 2 grooving, 3 boring bar, 4 drill ...
		std::wstring insertShapes;	//!< the insert shapes it takes ("TCDVKWE")
		double sideAngle = 0, endAngle = 0;	//!< as stored: drawing data, not the entering angle
		};
	HolderInfo LatheHolderInfo (long slot);

	// ---- ISO / ANSI CODES (no SDK) ------------------------------------------

	/// A holder style letter's ENTERING ANGLE (kr): the angle between the main
	/// cutting edge and the feed direction - 90 a square shoulder, no chip
	/// thinning; the chip is feed x sin(kr) thick. US shops say "lead angle" for
	/// 90 - kr. From ISO 5608:1995 table 3 (as published identically in IS
	/// 14863:2000), and style Q from ISO 6261 (boring bars, 107.5). The end-cutting
	/// styles (C, F, K, U, W, Y) are angles to a facing feed. 0 = no such style.
	inline double StyleAngle (wchar_t style)
		{
		switch (style)
			{
			case L'A': case L'C': case L'F': case L'G': return 90;
			case L'B': case L'K': case L'R': return 75;
			case L'D': case L'S': return 45;
			case L'E': case L'T': case L'W': return 60;
			case L'H': case L'Q': return 107.5;
			case L'J': case L'U': return 93;
			case L'L': return 95;
			case L'M': return 50;
			case L'N': return 63;
			case L'P': return 117.5;
			case L'V': return 72.5;
			case L'Y': return 85;
			default: return 0;
			}
		}

	/// The first ISO 5608 / ANSI holder code in some text, as its five letters -
	/// "PCLNR 2525M 12", "MCLNR-164D", "S25T-PCLNR12", Kennametal's "DTJNLS-164D"
	/// all give their PCLNR / MCLNR / DTJNL - or "". The letters: how the insert is
	/// held (C D M P S), the insert's shape, a style the table knows, the insert's
	/// clearance and the hand (R L N); then digits, after at most a maker's own
	/// letter and one dash or space. Clearances E and O are left out on purpose:
	/// "STEEL 4140" is not a holder.
	inline std::wstring HolderCode (const std::wstring &text)
		{
		std::wstring u;
		for (wchar_t c : text)
			u += static_cast<wchar_t> (std::towupper (c));
		auto in = [] (const wchar_t *set, wchar_t c)
			{
			return c != 0 && std::wstring (set).find (c) != std::wstring::npos;
			};
		for (size_t i = 0; i + 5 <= u.size (); ++i)
			{
			if (i > 0 && std::iswalnum (u[i - 1]))
				continue;
			if (!in (L"CDMPS", u[i]) || !in (L"ABCDEHKLMOPRSTVW", u[i + 1]) || StyleAngle (u[i + 2]) == 0
				|| !in (L"ABCDFGNP", u[i + 3]) || !in (L"RLN", u[i + 4]))
				continue;
			size_t k = i + 5;
			if (k < u.size () && std::iswalpha (u[k]))		// a maker's own sixth letter
				++k;
			if (k < u.size () && (u[k] == L'-' || u[k] == L' ' || u[k] == L'_'))
				++k;
			if (k < u.size () && std::iswdigit (u[k]))
				return u.substr (i, 5);
			}
		return std::wstring ();
		}

	/// The entering angle a holder code in some text gives (0: none, or a round
	/// insert's holder - a round's angle goes with the depth of cut).
	inline double HolderAngle (const std::wstring &text)
		{
		const std::wstring code = HolderCode (text);
		return code.empty () || code[1] == L'R' ? 0 : StyleAngle (code[2]);
		}

	/// An insert's IC (inscribed circle; a round insert's diameter) from the first
	/// ISO / ANSI insert code in some text, in mm (`mm`) or inches; 0 when none
	/// says. A metric code gives the cutting edge length in mm (CNMG 12 04 08,
	/// CNMG120408, DCMT 11T304), which the shape turns into its standard IC (12.70,
	/// 9.525); a round's is its diameter (RCMT 10T3MO: 10). An inch code gives the
	/// IC in eighths (CNMG 432, CNMG-432: 0.5; TPMT-21.51: 0.25). Shapes whose
	/// size cannot be told this way (K, parallelograms, polygons) give 0.
	inline double InsertCodeIC (const std::wstring &text, bool mm)
		{
		std::wstring u;
		for (wchar_t c : text)
			u += static_cast<wchar_t> (std::towupper (c));
		// IC per mm of cutting edge length, by shape (from the shape's geometry);
		// the IC is the standard one whose edge length the code gives.
		auto perLength = [] (wchar_t shape)
			{
			switch (shape)
				{
				case L'C': return 0.985;		// sin 80
				case L'D': return 0.819;		// sin 55
				case L'E': return 0.966;		// sin 75
				case L'M': return 0.998;		// sin 86
				case L'S': return 1.0;
				case L'T': return 0.577;		// 1 / sqrt 3
				case L'V': return 0.574;		// sin 35
				case L'W': return 1.588;
				default: return 0.0;
				}
			};
		const double standard[] = { 3.97, 4.76, 5.56, 6.35, 7.94, 9.525, 12.7, 15.875, 19.05, 25.4, 31.75 };
		for (size_t i = 0; i + 3 < u.size (); ++i)
			{
			if (i > 0 && std::iswalnum (u[i - 1]))
				continue;
			const wchar_t shape = u[i];
			if (std::wstring (L"CDEMSTVWR").find (shape) == std::wstring::npos
				|| std::wstring (L"ABCDEFGNPO").find (u[i + 1]) == std::wstring::npos || !std::iswalpha (u[i + 2]))
				continue;
			size_t k = i + 3;
			const bool four = k < u.size () && std::iswalpha (u[k]);	// the fourth letter (insert type)
			if (four)
				++k;
			// An old three-letter ANSI code (SNG-422, TPG-221) only with its dash.
			if (!four && (k >= u.size () || u[k] != L'-'))
				continue;
			if (k < u.size () && (u[k] == L'-' || u[k] == L' '))
				++k;
			size_t e = k;
			while (e < u.size () && (std::iswdigit (u[e]) || u[e] == L'.' || (e > k && u[e] == L'T')))
				++e;
			const std::wstring run = u.substr (k, e - k);
			if (run.empty () || !std::iswdigit (run[0]))
				continue;
			const bool twoDigits = run.size () >= 2 && std::iswdigit (run[1]);
			const bool spaced = run.size () == 2 && e + 1 < u.size () && u[e] == L' '
								&& (std::iswdigit (u[e + 1]) || u[e + 1] == L'T');
			const bool metric = twoDigits && run.find (L'.') == std::wstring::npos && (run.size () >= 4 || spaced);
			double icMm = 0;
			if (metric)
				{
				const double size = (run[0] - L'0') * 10.0 + (run[1] - L'0');
				if (shape == L'R')
					icMm = size;
				else if (perLength (shape) > 0)
					{
					// The code's figure is the edge length cut down to whole mm (C 12.7 IC:
					// 12.9 -> 12; W 9.525: 5.998 -> 06, hence the hair of slack).
					for (double s : standard)
						{
						const double edge = s / perLength (shape);
						if (size <= edge + 0.1 && edge < size + 1)
							{
							icMm = s;
							break;
							}
						}
					}
				}
			else if (run.size () == 3 || (run.size () > 3 && run.find (L'.') != std::wstring::npos))
				{
				// IC, thickness and radius figures (432, 21.51) - one figure alone is
				// a word and a number ("TEST 1"), not a code.
				if (run[0] == L'0' || (shape != L'R' && perLength (shape) == 0))
					continue;
				const double inches = (run[0] - L'0') / 8.0;
				return mm ? inches * 25.4 : inches;
				}
			if (icMm > 0)
				return mm ? icMm : std::round (icMm / 25.4 * 10000.0) / 10000.0;
			}
		return 0;
		}
	}
