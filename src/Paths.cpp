#include "stdafx.h"
#include "MastercamSdk.h"
#include "Paths.h"
#include "Csv.h"
#include "BnciReadWrite_CH.h"

#include <map>

#include <algorithm>
#include <cmath>

namespace
	{
	const double kPi = 3.14159265358979323846;

	/// THE SPINDLE AS THE NCI SETS IT. Each section of an operation's NCI opens
	/// with a 1002 line giving the speed - positive RPM, negative surface speed
	/// (SFM, or m/min in a metric part) - and the cap. A groove's finish passes,
	/// for one, run at their own speed in their own section.
	struct Spindle
		{
		double speed = 0;		//!< RPM, or surface speed when css
		bool css = false;
		double cap = 0;			//!< max RPM; 0 = none
		bool mm = false;
		};

	/// RPM at a radius.
	double RpmAt (const Spindle &sp, double radius)
		{
		if (!sp.css)
			return sp.speed;
		const double dia = 2.0 * std::fabs (radius);
		// SFM -> RPM: 12 * SFM / (pi * D in inches); m/min -> RPM: 1000 * m/min / (pi * D in mm).
		const double k = sp.mm ? 1000.0 : 12.0;
		double rpm = dia > 1e-9 ? k * sp.speed / (kPi * dia) : 1e12;
		if (sp.cap > 0 && rpm > sp.cap)
			rpm = sp.cap;
		return rpm;
		}

	/// Seconds for one feed move of `len` at radius `r`. The MOVE'S OWN feed
	/// sign says per rev (negative) or per minute (positive).
	double MoveSeconds (const Spindle &sp, double len, double r, double feed)
		{
		const double f = std::fabs (feed);
		if (f <= 0 || len <= 0)
			return 0;
		if (feed < 0)
			{
			const double rpm = RpmAt (sp, r);
			return rpm > 0 ? 60.0 * len / (f * rpm) : 0;
			}
		return 60.0 * len / f;
		}
	}

namespace Paths
	{
	Totals Walk (operation &op)
		{
		Totals t;
		CBnciReadWrite nci;
		const INT_PTR n = nci.ReadSection (op.op_idn, true);
		if (n <= 0)
			return t;

		bool have = false;
		double at[3] = { 0, 0, 0 };
		auto visit = [&t, &have] (const double *p)
			{
			for (int a = 0; a < 3; ++a)
				{
				if (!have || p[a] < t.min[a]) t.min[a] = p[a];
				if (!have || p[a] > t.max[a]) t.max[a] = p[a];
				}
			have = true;
			};

		Spindle sp;
		sp.speed = std::fabs (static_cast<double> (op.tl.rpm));
		sp.css = op.tl.use_css;
		sp.cap = std::fabs (static_cast<double> (op.tl.max_ss));
		sp.mm = op.tl.mm;

		// Pieces of feed moves by (feed, section): (length, radius).
		struct Key
			{
			double feed;
			long section;
			int axis;
			bool operator< (const Key &o) const
				{
				if (section != o.section) return section < o.section;
				if (feed != o.feed) return feed < o.feed;
				return axis < o.axis;
				}
			};
		// A lathe move's direction: 1 mostly along Z (axial), 2 mostly along X.
		auto axisOf = [] (bool lathe, const double *from, const double *to)
			{
			if (!lathe)
				return 0;
			return std::fabs (to[2] - from[2]) >= std::fabs (to[0] - from[0]) ? 1 : 2;
			};
		std::map<Key, std::vector<std::pair<double, double>>> bucketOf;
		std::map<long, Spindle> spindleOf;
		spindleOf[0] = sp;

		bool started = false;
		for (INT_PTR i = 0; i < n; ++i)
			{
			const nci_bin *b = nci[i];
			if (b == nullptr)
				continue;
			const NCI_GCODE g = b->gcode;
			if (static_cast<int> (g) == 1002 && b->type == NCI_LATHE_TYPE)
				{
				const long rpm = b->u.l1002.rpm;
				sp.css = rpm < 0;
				sp.speed = std::fabs (static_cast<double> (rpm));
				sp.cap = std::fabs (b->u.l1002.max_rpm);
				++t.sections;
				spindleOf[t.sections] = sp;
				continue;
				}
			// CANNED DRILLING: one line per hole, no moves. Each hole feeds from the
			// feed plane down to the depth - taken from the OPERATION's own planes
			// (a real part: feed plane 5.25, depth 2.75, 5 in/min = Mastercam's 30 s
			// exactly), the NCI line's heights only when the planes give nothing.
			if (g == NCI_DRILL_START || g == NCI_DRILL_HOLE)
				{
				const bool lathe = b->type == NCI_LATHE_TYPE;
				double len = 0, feed = 0, r = 0;
				const double top = op.cmn.top_stock;
				const double feedPlane = op.cmn.feed_inc ? top + op.cmn.feed_pln : op.cmn.feed_pln;
				const double bottom = op.cmn.depth_inc ? top + op.cmn.depth : op.cmn.depth;
				const double fromPlanes = feedPlane - bottom;
				if (g == NCI_DRILL_START && lathe)
					{
					len = std::fabs (b->u.l81.ref_hgt) + std::fabs (b->u.l81.xdrl);
					feed = b->u.l81.feed;
					r = b->u.l81.ep1[0];
					}
				else if (g == NCI_DRILL_START)
					{
					len = std::fabs (b->u.m81.ref_hgt) + std::fabs (b->u.m81.z);
					feed = b->u.m81.feed;
					}
				else
					{
					len = std::fabs (b->u.m100.z) + std::fabs (b->u.m100.depth);
					feed = b->u.m100.feed;
					}
				if (fromPlanes > 0)
					len = fromPlanes;
				if (len > 0 && feed != 0)
					{
					++t.holes;
					t.drillLength += len;
					t.cutLength += len;
					t.feedSeconds += MoveSeconds (sp, len, r, feed);
					bucketOf[{ feed, t.sections, 0 }].push_back ({ len, r });
					}
				// The hole's bottom is part of where the tool goes (a mill hole: Z down
				// from where the tool stands to the depth - Mastercam's extents agree).
				if (!lathe && fromPlanes > 0 && have)
					{
					const double p[3] = { at[0], at[1], bottom };
					visit (p);
					}
				continue;
				}
			if (g != NCI_RAPID && g != NCI_LINEAR && g != NCI_ARC_CW && g != NCI_ARC_CCW)
				continue;

			// WHERE A LATHE MOVE KEEPS ITS POINT - read off a real part's NCI, where
			// this exact reading reproduces backplot's feed length to 4 decimals:
			//   rapid / line:  (X, 0, Z)
			//   arc:           end (X, Z, 0) and centre (X, Z, 0) - Z in the SECOND slot.
			// Everything below works on (X, 0, Z).
			const bool arc = g == NCI_ARC_CW || g == NCI_ARC_CCW;
			const bool lathe = b->type == NCI_LATHE_TYPE;
			const p_3d &ep = arc ? b->u.l2.ep1 : b->u.l1.ep1;
			const double end[3] = { ep[0], arc && lathe ? 0.0 : ep[1], arc && lathe ? ep[1] : ep[2] };
			++t.moves;
			if (!started)
				{
				// The first move's start is wherever the tool was before this
				// operation; its length is not counted.
				std::copy (end, end + 3, at);
				started = true;
				visit (end);
				continue;
				}

			if (g == NCI_RAPID || g == NCI_LINEAR)
				{
				const double len = std::sqrt ((end[0] - at[0]) * (end[0] - at[0])
											  + (end[1] - at[1]) * (end[1] - at[1])
											  + (end[2] - at[2]) * (end[2] - at[2]));
				if (g == NCI_RAPID)
					t.rapidLength += len;
				else
					{
					t.cutLength += len;
					// A move across diameters (a facing pass) in short pieces: under CSS
					// its speed changes along it.
					const double dx = std::fabs (end[0] - at[0]);
					const int k = (std::max) (1, (std::min) (64, static_cast<int> (std::ceil (dx / 0.05))));
					auto &bucket = bucketOf[{ b->u.l1.feed, t.sections, axisOf (lathe, at, end) }];
					for (int q = 0; q < k; ++q)
						{
						const double r = at[0] + (end[0] - at[0]) * (q + 0.5) / k;
						t.feedSeconds += MoveSeconds (sp, len / k, r, b->u.l1.feed);
						bucket.push_back ({ len / k, r });
						}
					}
				}
			else
				{
				// A lathe arc lies in X-Z (plane code 2), swept from Z towards X -
				// confirmed: no arc on the real part comes out over 180 degrees.
				const short pln = b->u.l2.pln;
				if (pln >= 0 && pln <= 2)
					++t.planes[pln];
				const int ua = lathe ? 2 : pln == 1 ? 1 : pln == 2 ? 2 : 0;
				const int va = lathe ? 0 : pln == 1 ? 2 : pln == 2 ? 0 : 1;
				const p_3d &cp = b->u.l2.cpt;
				const double c[3] = { cp[0], lathe ? 0.0 : cp[1], lathe ? cp[1] : cp[2] };
				const double su = at[ua] - c[ua], sv = at[va] - c[va];
				const double eu = end[ua] - c[ua], ev = end[va] - c[va];
				const double r = std::sqrt (su * su + sv * sv);
				double sweep = std::atan2 (ev, eu) - std::atan2 (sv, su);	// anticlockwise +
				if (g == NCI_ARC_CW)
					sweep = -sweep;
				while (sweep <= 1e-12)
					sweep += 2.0 * kPi;
				const double len = r * sweep;
				t.cutLength += len;
				t.arcLength += len;
				++t.arcs;
				if (sweep > kPi + 1e-6)
					++t.arcsOverHalf;

				// Time along the arc, in short pieces - its radius from the spindle
				// changes as it goes.
				const int pieces = 16;
				auto &piecesOfArc = bucketOf[{ b->u.l2.feed, t.sections, axisOf (lathe, at, end) }];
				const double a0 = std::atan2 (sv, su), dir = g == NCI_ARC_CW ? -1.0 : 1.0;
				for (int k = 0; k < pieces; ++k)
					{
					const double a = a0 + dir * sweep * (k + 0.5) / pieces;
					double p[3] = { c[0], c[1], c[2] };
					p[ua] += r * std::cos (a);
					p[va] += r * std::sin (a);
					visit (p);
					t.feedSeconds += MoveSeconds (sp, len / pieces, p[0], b->u.l2.feed);
					piecesOfArc.push_back ({ len / pieces, p[0] });
					}
				}
			std::copy (end, end + 3, at);
			visit (end);
			}
		// The groups: each (feed, section) with that section's spindle, banded.
		for (const auto &kv : bucketOf)
			{
			Estimate::Group g;
			g.feed = kv.first.feed;
			const Spindle &s2 = spindleOf.count (kv.first.section) ? spindleOf[kv.first.section] : sp;
			g.speed = s2.speed;
			g.css = s2.css;
			g.cap = s2.cap;
			g.mm = s2.mm;
			g.axis = kv.first.axis;
			g.bands = Estimate::Bands (kv.second, 6);
			t.groups.push_back (g);
			}
		t.ok = t.moves > 0 || t.holes > 0;
		return t;
		}

	std::wstring Listing (operation &op)
		{
		CBnciReadWrite nci;
		const INT_PTR n = nci.ReadSection (op.op_idn, true);
		std::wstring s;
		auto d = [] (double v) { wchar_t b[40]; swprintf_s (b, L"%.6f", v); return std::wstring (b); };
		for (INT_PTR i = 0; i < n; ++i)
			{
			const nci_bin *b = nci[i];
			if (b == nullptr)
				continue;
			s += std::to_wstring (op.op_idn) + L"," + std::to_wstring (i) + L","
				 + std::to_wstring (static_cast<int> (b->gcode)) + L"," + std::to_wstring (static_cast<int> (b->type));
			const NCI_GCODE g = b->gcode;
			if (g == NCI_RAPID || g == NCI_LINEAR)
				s += L",," + d (b->u.l1.ep1[0]) + L"," + d (b->u.l1.ep1[1]) + L"," + d (b->u.l1.ep1[2])
					 + L",,,," + d (b->u.l1.feed) + L"," + std::to_wstring (b->u.l1.ccomp) + L","
					 + std::to_wstring (b->u.l1.position);
			else if (g == NCI_ARC_CW || g == NCI_ARC_CCW)
				s += L"," + std::to_wstring (b->u.l2.pln) + L"," + d (b->u.l2.ep1[0]) + L","
					 + d (b->u.l2.ep1[1]) + L"," + d (b->u.l2.ep1[2]) + L"," + d (b->u.l2.cpt[0]) + L","
					 + d (b->u.l2.cpt[1]) + L"," + d (b->u.l2.cpt[2]) + L"," + d (b->u.l2.feed) + L","
					 + std::to_wstring (b->u.l2.ccomp) + L"," + std::to_wstring (b->u.l2.position);
			else if (static_cast<int> (g) == 1002 && b->type == NCI_LATHE_TYPE)
				s += L",,,,,,,," + d (b->u.l1002.feed) + L",rpm " + std::to_wstring (b->u.l1002.rpm)
					 + L",max " + d (b->u.l1002.max_rpm);
			s += L"\r\n";
			}
		return s;
		}

	std::wstring Describe (const operation &op, const Totals &t, double mastercamSeconds)
		{
		auto n = [] (double v) { return Csv::Tidy (std::round (v * 1000.0) / 1000.0); };
		auto hms = [] (double s)
			{
			const long long v = static_cast<long long> (s + 0.5);
			wchar_t buf[32];
			swprintf_s (buf, L"%lld:%02lld:%02lld", v / 3600, (v / 60) % 60, v % 60);
			return std::wstring (buf);
			};
		std::wstring s = L"path op " + std::to_wstring (op.op_idn) + L": ";
		if (!t.ok)
			return s + L"no NCI to walk";
		s += std::to_wstring (t.sections) + L" sections, " + std::to_wstring (t.moves) + L" moves, "
			 + (t.holes ? std::to_wstring (t.holes) + L" holes (drill feed " + n (t.drillLength) + L"), " : L"")
			 + L"cut " + n (t.cutLength) + L" (" + std::to_wstring (t.arcs)
			 + L" arcs " + n (t.arcLength) + L", " + std::to_wstring (t.arcsOverHalf) + L" over 180deg, planes "
			 + std::to_wstring (t.planes[0]) + L"/" + std::to_wstring (t.planes[1]) + L"/"
			 + std::to_wstring (t.planes[2]) + L"), rapid " + n (t.rapidLength)
			 + L" | X " + n (t.min[0]) + L".." + n (t.max[0])
			 + L" Y " + n (t.min[1]) + L".." + n (t.max[1])
			 + L" Z " + n (t.min[2]) + L".." + n (t.max[2])
			 + L" | feed time " + hms (t.feedSeconds) + L" vs Mastercam " + hms (mastercamSeconds);
		return s;
		}
	}
