#include "stdafx.h"
#include "MastercamSdk.h"
#include "StockSim.h"
#include "StockRaster.h"
#include "Paths.h"
#include "Util.h"
#include "BnciReadWrite_CH.h"
#include "TlServices_CH.h"
#include "TlMgr_CH.h"
#include "ILTool_CH.h"
#include "TpVars_CH.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdint>
#include <map>
#include <set>
#include <vector>

namespace
	{
	using namespace StockRaster;

	bool gLog = false;			// the probe's log lines (diag bit 4)
	void LogIf (const std::filesystem::path &part, const std::wstring &line)
		{
		if (gLog)
			Util::Log (part, line);
		}

	std::wstring F (double v, int places = 4)
		{
		wchar_t b[48];
		swprintf_s (b, L"%.*f", places, v);
		return b;
		}

	/// THE TOOL'S SHAPE: the insert's cutting boundary as Mastercam draws it
	/// ("on the XY plane": read here as X -> Z, Y -> X radius, the control point
	/// at the origin), its convex hull, and the nose circle alone for comparison.
	struct Shape
		{
		std::vector<P> insert;		//!< hull of the cut boundary
		std::vector<P> nose;		//!< the corner-radius circle only
		bool fromBoundary = false;
		};

	Shape ShapeOf (long slot, const std::filesystem::path &part)
		{
		Shape s;
		std::wstring line = L"stock sim tool slot " + std::to_wstring (slot) + L": ";
		Cnc::Tool::TlMgr *mgr = Cnc::Tool::GetTlMgr ();
		Cnc::Tool::ILToolCPtr tool;
		double rc = 0;
		std::vector<P> pts;
		std::vector<std::pair<P, double>> arcs;		// centre, radius
		if (mgr != nullptr && mgr->Find (slot, tool) && tool)
			{
			rc = tool->GetInsertCornerRadius ();
			const p_2d tc = tool->GetToolCenter ();
			line += L"type " + std::to_wstring (static_cast<int> (tool->GetType ()))
					+ L" insert \"" + std::wstring (tool->GetInsertName ().GetString ()) + L"\" rc " + F (rc)
					+ L" comp " + std::to_wstring (static_cast<int> (tool->GetCompMethod ()))
					+ L" tpOrient " + std::to_wstring (static_cast<int> (tool->GetToolpathOrientation ()))
					+ L" holderOrient " + std::to_wstring (static_cast<int> (tool->GetHolderOrientation ()))
					+ L" quadrant " + std::to_wstring (static_cast<int> (tool->GetCompQuadrant ()))
					+ L" rh " + std::to_wstring (tool->GetIsRightHanded () ? 1 : 0)
					+ L" vert " + std::to_wstring (tool->GetIsVertical () ? 1 : 0)
					+ L" rev " + std::to_wstring (tool->GetIsReversed () ? 1 : 0)
					+ L" angle " + F (tool->GetToolAngle (), 2)
					+ L" center (" + F (tc[0]) + L"," + F (tc[1]) + L")";
			const std::vector<tp_ent> cut = tool->GetCutBoundary ();
			line += L" | cut boundary " + std::to_wstring (cut.size ()) + L" ents:";
			int shown = 0;
			for (const tp_ent &e : cut)
				{
				std::wstring d;
				if (e.id == L_ID)
					{
					pts.push_back ({ e.u.li.e1[0], e.u.li.e1[1] });
					pts.push_back ({ e.u.li.e2[0], e.u.li.e2[1] });
					d = L" L(" + F (e.u.li.e1[0]) + L"," + F (e.u.li.e1[1]) + L")-(" + F (e.u.li.e2[0]) + L","
						+ F (e.u.li.e2[1]) + L")";
					}
				else if (e.id == A_ID)
					{
					const a_3d &a = e.u.ar;
					const int n = (std::max) (4, static_cast<int> (std::ceil (std::fabs (a.sw) / (kPi / 36))));
					for (int k = 0; k <= n; ++k)
						{
						const double t = a.sa + a.sw * k / n;
						pts.push_back ({ a.c[0] + a.r * std::cos (t), a.c[1] + a.r * std::sin (t) });
						}
					// The stored end points too, in case the angles are in another sense.
					pts.push_back ({ a.ep1[0], a.ep1[1] });
					pts.push_back ({ a.ep2[0], a.ep2[1] });
					arcs.push_back ({ { a.c[0], a.c[1] }, a.r });
					d = L" A(c " + F (a.c[0]) + L"," + F (a.c[1]) + L" r " + F (a.r) + L" sa " + F (a.sa * 180 / kPi, 1)
						+ L" sw " + F (a.sw * 180 / kPi, 1) + L" ep " + F (a.ep1[0]) + L"," + F (a.ep1[1]) + L" "
						+ F (a.ep2[0]) + L"," + F (a.ep2[1]) + L" view " + std::to_wstring (a.view) + L")";
					}
				else
					d = L" id" + std::to_wstring (e.id);
				d += L"[src" + std::to_wstring (e.source) + L" lvl" + std::to_wstring (e.level) + L"]";
				if (shown++ < 24)
					line += d;
				}
			line += L" | non-cut " + std::to_wstring (tool->GetNonCutBoundary ().size ()) + L" ents";
			}
		else
			line += L"no lathe tool";

		if (pts.size () >= 3)
			{
			s.insert = Hull (pts);
			s.fromBoundary = s.insert.size () >= 3;
			}
		// The nose: the arc of the corner radius nearest the control point.
		P nc { 0, 0 };
		double nr = rc;
		{
		double best = 1e300;
		for (const auto &a : arcs)
			{
			if (rc > 0 && std::fabs (a.second - rc) > 1e-3)
				continue;
			const double d = std::hypot (a.first.z, a.first.x);
			if (d < best)
				{
				best = d;
				nc = a.first;
				nr = a.second;
				}
			}
		}
		if (nr <= 0)
			nr = 0.002;
		for (int k = 0; k < 32; ++k)
			s.nose.push_back ({ nc.z + nr * std::cos (2 * kPi * k / 32), nc.x + nr * std::sin (2 * kPi * k / 32) });
		if (!s.fromBoundary)
			s.insert = s.nose;

		double z0 = 1e300, z1 = -1e300, x0 = 1e300, x1 = -1e300;
		for (const P &p : s.insert)
			{
			z0 = (std::min) (z0, p.z); z1 = (std::max) (z1, p.z);
			x0 = (std::min) (x0, p.x); x1 = (std::max) (x1, p.x);
			}
		line += L" | hull " + std::to_wstring (s.insert.size ()) + L" pts, X(=bdy x) " + F (z0) + L".." + F (z1)
				+ L" Y(=bdy y) " + F (x0) + L".." + F (x1) + L", control point to hull "
				+ F (SignedDistance (s.insert, { 0, 0 })) + L", nose centre (" + F (nc.z) + L"," + F (nc.x) + L") r "
				+ F (nr) + (s.fromBoundary ? L"" : L" [NO BOUNDARY - nose circle used for both]");
		LogIf (part, line);
		return s;
		}

	/// One straight piece of the path (arcs come as chords), control point.
	struct Seg
		{
		P a, b;
		bool feed;
		double rate = 0;		//!< the move's feed: negative per rev, positive per minute
		double speed = 0;		//!< the section's spindle: RPM, or surface speed when css
		bool css = false;
		double cap = 0;
		};

	/// The boundary's entities, copied out under a guard: bdryEnts is only good
	/// while the boundary is in memory, and nothing in the SDK says when that is.
	int CopyBoundary (const lathe_bdry *b, tp_ent *out, int max)
		{
		__try
			{
			if (b->bdryEnts == nullptr || b->nRec <= 0)
				return 0;
			const int n = b->nRec < max ? static_cast<int> (b->nRec) : max;
			memcpy (out, b->bdryEnts, sizeof (tp_ent) * n);
			return n;
			}
		__except (EXCEPTION_EXECUTE_HANDLER)
			{
			return -1;
			}
		}

	/// A LATHE BOUNDARY ENTITY (LBDRY: the stock outline the machine group's
	/// stock setup draws, and each op's stock left behind) as a closed outline
	/// in its own XY, not yet mapped onto (z, radius).
	struct Bdry
		{
		bool ok = false;
		std::vector<P> raw;		//!< (bdy x, bdy y)
		std::wstring info;
		};

	Bdry ReadBoundary (long id)
		{
		Bdry r;
		r.info = L"ent " + std::to_wstring (id);
		if (id <= 0)
			return r;
		ent e;
		bool got = false;
		GetEntityByID (id, e, &got);
		if (!got)
			{
			r.info += L" not found";
			return r;
			}
		const lathe_bdry &b = e.u.lbdry;
		r.info += L" type " + std::to_wstring (e.id) + L"/assoc " + std::to_wstring (static_cast<int> (e.assoc_id))
				  + L" nRec " + std::to_wstring (b.nRec) + L" kind " + std::to_wstring (b.type) + L" valid "
				  + std::to_wstring (b.valid) + L" op " + std::to_wstring (b.op_idn) + L" xlate (" + F (b.xlateVec[0])
				  + L"," + F (b.xlateVec[1]) + L") mirrored " + std::to_wstring (b.mirrored) + L" transferred "
				  + std::to_wstring (b.transferred) + L" scale " + F (b.scale) + L" view " + std::to_wstring (b.view_n)
				  + L" depth " + F (b.rel_depth) + L" ram " + std::to_wstring (b.ram) + L" m [";
		for (int i = 0; i < 3; ++i)
			r.info += F (b.view_m[i][0], 3) + L"," + F (b.view_m[i][1], 3) + L","
					  + F (b.view_m[i][2], 3) + (i < 2 ? L"; " : L"]");
		if (static_cast<int> (e.assoc_id) != static_cast<int> (LBDRY_ID))
			{
			r.info += L" NOT A LATHE BOUNDARY";
			return r;
			}
		std::vector<tp_ent> ents (20000);
		const int n = CopyBoundary (&b, ents.data (), static_cast<int> (ents.size ()));
		if (n < 0)
			{
			r.info += L" ENTITIES UNREADABLE";
			return r;
			}
		std::vector<std::vector<P>> pieces;
		int shown = 0;
		for (int k = 0; k < n; ++k)
			{
			const tp_ent &t = ents[k];
			std::vector<P> pc;
			if (t.id == L_ID)
				{
				pc = { { t.u.li.e1[0], t.u.li.e1[1] }, { t.u.li.e2[0], t.u.li.e2[1] } };
				if (shown++ < 12)
					r.info += L" L(" + F (pc[0].z) + L"," + F (pc[0].x) + L")-(" + F (pc[1].z) + L"," + F (pc[1].x) + L")";
				}
			else if (t.id == A_ID)
				{
				const a_3d &a = t.u.ar;
				const int m = (std::max) (2, static_cast<int> (std::ceil (std::fabs (a.sw) / (kPi / 180))));
				for (int q = 0; q <= m; ++q)
					{
					const double ang = a.sa + a.sw * q / m;
					pc.push_back ({ a.c[0] + a.r * std::cos (ang), a.c[1] + a.r * std::sin (ang) });
					}
				// Ends as stored, in case the angles run the other way.
				const double d1 = std::hypot (pc.front ().z - a.ep1[0], pc.front ().x - a.ep1[1]);
				const double d2 = std::hypot (pc.back ().z - a.ep2[0], pc.back ().x - a.ep2[1]);
				if (shown++ < 12)
					r.info += L" A(c " + F (a.c[0]) + L"," + F (a.c[1]) + L" r " + F (a.r) + L" ends off " + F (d1)
							  + L"/" + F (d2) + L")";
				}
			if (!pc.empty ())
				pieces.push_back (pc);
			}
		double gap = 0;
		r.raw = Chain (pieces, gap);
		// The entities are stored times the scale (1000 on a real part: a 7.0 radius as 7000).
		if (b.scale > 0)
			{
			for (P &p : r.raw)
				{
				p.z /= b.scale;
				p.x /= b.scale;
				}
			gap /= b.scale;
			}
		double x0 = 1e300, x1 = -1e300, y0 = 1e300, y1 = -1e300;
		for (const P &p : r.raw)
			{
			x0 = (std::min) (x0, p.z); x1 = (std::max) (x1, p.z);
			y0 = (std::min) (y0, p.x); y1 = (std::max) (y1, p.x);
			}
		r.info += L" | " + std::to_wstring (n) + L" ents chained, worst gap " + F (gap) + L", bdy x " + F (x0) + L".."
				  + F (x1) + L" y " + F (y0) + L".." + F (y1);
		r.ok = r.raw.size () >= 3 && gap < 0.01;
		if (!r.ok)
			r.info += L" - NOT A CLOSED OUTLINE";
		return r;
		}

	/// Boundary XY onto the lathe half-plane: swap = bdy y is Z; dia = the
	/// radial coordinate is a diameter.
	std::vector<P> Mapped (const std::vector<P> &raw, bool swap, bool dia)
		{
		std::vector<P> out;
		out.reserve (raw.size ());
		for (const P &p : raw)
			{
			P q = swap ? P { p.x, p.z } : p;
			q.x = std::fabs (q.x) * (dia ? 0.5 : 1.0);
			out.push_back (q);
			}
		return out;
		}

	struct OpPath
		{
		operation *op = nullptr;
		std::vector<Seg> segs;
		long holes = 0;			//!< canned drill holes - not simulated
		long otherType = 0;		//!< moves that are not lathe NCI (mill moves in an MT group)
		bool afterFlip = false;	//!< after a stock flip / transfer - in the wrong frame
		};

	OpPath Read (operation &op)
		{
		OpPath r;
		r.op = &op;
		CBnciReadWrite nci;
		const INT_PTR n = nci.ReadSection (op.op_idn, true);
		bool started = false;
		P at { 0, 0 };
		// The spindle as the NCI sets it (see Paths.cpp): the op's own, then each 1002 line's.
		double speed = std::fabs (static_cast<double> (op.tl.rpm)), cap = std::fabs (static_cast<double> (op.tl.max_ss));
		bool css = op.tl.use_css;
		for (INT_PTR i = 0; i < n; ++i)
			{
			const nci_bin *b = nci[i];
			if (b == nullptr)
				continue;
			const NCI_GCODE g = b->gcode;
			if (static_cast<int> (g) == 1002 && b->type == NCI_LATHE_TYPE)
				{
				css = b->u.l1002.rpm < 0;
				speed = std::fabs (static_cast<double> (b->u.l1002.rpm));
				cap = std::fabs (b->u.l1002.max_rpm);
				continue;
				}
			if (g == NCI_DRILL_START || g == NCI_DRILL_HOLE)
				{
				++r.holes;
				continue;
				}
			if (g != NCI_RAPID && g != NCI_LINEAR && g != NCI_ARC_CW && g != NCI_ARC_CCW)
				continue;
			if (b->type != NCI_LATHE_TYPE)
				{
				++r.otherType;
				continue;
				}
			// Lathe NCI (see Paths.cpp): line/rapid end (X, 0, Z); arc end and
			// centre (X, Z, 0); X a radius.
			const bool arc = g == NCI_ARC_CW || g == NCI_ARC_CCW;
			const p_3d &ep = arc ? b->u.l2.ep1 : b->u.l1.ep1;
			const P end { arc ? ep[1] : ep[2], ep[0] };
			if (!started)
				{
				at = end;
				started = true;
				continue;
				}
			if (!arc)
				r.segs.push_back ({ at, end, g == NCI_LINEAR, b->u.l1.feed, speed, css, cap });
			else
				{
				P prev = at;
				for (const P &p : Chords (at, end, { b->u.l2.cpt[1], b->u.l2.cpt[0] }, g == NCI_ARC_CW, 0.0005))
					{
					r.segs.push_back ({ prev, p, true, b->u.l2.feed, speed, css, cap });
					prev = p;
					}
				}
			at = end;
			}
		return r;
		}

	std::wstring Ms (double s)
		{
		const long long v = static_cast<long long> (s + 0.5);
		wchar_t buf[32];
		swprintf_s (buf, L"%lld:%02lld", v / 60, v % 60);
		return buf;
		}
	}

namespace StockSim
	{
	std::map<long, Result> Run (const std::filesystem::path &part, bool log)
		{
		gLog = log;
		std::map<long, Result> results;
		const auto t0 = std::chrono::steady_clock::now ();
		// Machine groups by id, and each op's root (machine) group.
		std::map<long, op_group *> groups;
		// GetAt is marked deprecated (for OpGroupsAPI, in a library this add-in does
		// not link) - it still works.
		TpGrpList &grpList = TpMainGrpMgr.GetMainGrpList ();
#pragma warning(push)
#pragma warning(disable : 4996)
		for (INT_PTR i = 0; i < static_cast<INT_PTR> (grpList.GetSize ()); ++i)
			if (op_group *g = grpList.GetAt (i); g != nullptr && g->grp_idn > 0)
				groups[g->grp_idn] = g;
#pragma warning(pop)
		auto rootOf = [&groups] (long id) -> op_group *
			{
			op_group *g = nullptr;
			for (int guard = 0; guard < 64; ++guard)
				{
				const auto it = groups.find (id);
				if (it == groups.end ())
					return g;
				g = it->second;
				if (g->parent_grp_idn <= 0)
					return g;
				id = g->parent_grp_idn;
				}
			return g;
			};

		Cnc::Tool::TpPartOpList &opList = TpMainOpMgr.GetMainOpList ();
		std::vector<std::pair<op_group *, std::vector<operation *>>> byGroup;
		for (INT_PTR i = 0; i < opList.GetSize (); ++i)
			{
			operation *op = opList.GetAt (i);
			if (op == nullptr)
				continue;
			op_group *g = rootOf (op->cmn.grp_idn);
			auto it = std::find_if (byGroup.begin (), byGroup.end (), [g] (const auto &p) { return p.first == g; });
			if (it == byGroup.end ())
				{
				byGroup.push_back ({ g, {} });
				it = byGroup.end () - 1;
				}
			it->second.push_back (op);
			}

		std::map<long, Shape> shapes;
		for (auto &grp : byGroup)
			{
			op_group *g = grp.first;
			std::wstring head = L"stock sim group ";
			if (g == nullptr)
				{
				LogIf (part, head + L"(none found) - " + std::to_wstring (grp.second.size ()) + L" ops skipped");
				continue;
				}
			const group_pg3 &s3 = g->ogi.pg3;
			head += std::to_wstring (g->grp_idn) + L" \"" + std::wstring (g->name) + L"\" product "
					+ std::to_wstring (static_cast<int> (g->product)) + L" jobSetupEnt "
					+ std::to_wstring (g->jobSetupDataEntityIdn) + L" | pg3 x " + F (s3.x) + L" y " + F (s3.y)
					+ L" z " + F (s3.z) + L" shape " + std::to_wstring (s3.stock_shape) + L" lstock_id "
					+ std::to_wstring (s3.lstock_id[0]) + L"," + std::to_wstring (s3.lstock_id[1]) + L" rstock_id "
					+ std::to_wstring (s3.rstock_id[0]) + L"," + std::to_wstring (s3.rstock_id[1]);
			for (int k = 0; k < 2; ++k)
				{
				const barstock_type &b = s3.stock_def[k];
				head += std::wstring (k == 0 ? L" | LEFT" : L" | RIGHT") + L" od " + F (b.od) + L" id " + F (b.id)
						+ L" len " + F (b.length) + L" refZ " + F (b.refZ) + L" refAtMaxZ " + std::to_wstring (b.refAtMaxZ)
						+ L" hole " + std::to_wstring (b.holeInStock) + L" margins " + F (b.margins[0]) + L","
						+ F (b.margins[1]) + L"," + F (b.margins[2]) + L"," + F (b.margins[3])
						+ L" useMargins " + std::to_wstring (b.useMargins) + L" show " + std::to_wstring (s3.show_stock[k]);
				}
			LogIf (part, head);
			if (g->product != PRODUCT_LATHE && g->product != PRODUCT_MT)
				continue;

			// The paths, and the tools' shapes.
			std::vector<OpPath> paths;
			bool flipped = false;
			for (operation *op : grp.second)
				{
				const long code = static_cast<long> (op->opcode);
				if (code == TP_LSTOCK_FLIP || code == TP_LSTOCK_XFER)
					flipped = true;
				if (code == TP_LSTOCK_FLIP || code == TP_LSTOCK_XFER)
					LogIf (part, L"stock sim op " + std::to_wstring (op->op_idn) + L": STOCK "
									 + (code == TP_LSTOCK_FLIP ? L"FLIP" : L"TRANSFER")
									 + L" - not followed; what comes after is in the wrong frame");
				if (op->db.nci_flag)
					{
					LogIf (part, L"stock sim op " + std::to_wstring (op->op_idn) + L": needs regen - skipped");
					continue;
					}
				OpPath p = Read (*op);
				p.afterFlip = flipped;
				if (p.segs.empty () && p.holes == 0)
					continue;
				if (!shapes.count (op->tl.slot))
					shapes[op->tl.slot] = ShapeOf (op->tl.slot, part);
				paths.push_back (std::move (p));
				}
			if (paths.empty ())
				continue;

			// The stock: the left spindle's bar, else the right's.
			const barstock_type *bar = s3.stock_def[0].od > 0 ? &s3.stock_def[0]
									   : s3.stock_def[1].od > 0 ? &s3.stock_def[1] : nullptr;
			double odR = 0, idR = 0, zA = 0, zB = 0;
			// Where the feed moves go (with the tool around them).
			double fz0 = 1e300, fz1 = -1e300, fx1 = 0;
			for (const OpPath &p : paths)
				{
				const Shape &sh = shapes[p.op->tl.slot];
				double hz0 = 0, hz1 = 0, hx = 0;
				for (const P &q : sh.insert)
					{
					hz0 = (std::min) (hz0, q.z);
					hz1 = (std::max) (hz1, q.z);
					hx = (std::max) (hx, std::fabs (q.x));
					}
				for (const Seg &s : p.segs)
					if (s.feed)
						for (const P &q : { s.a, s.b })
							{
							fz0 = (std::min) (fz0, q.z + hz0);
							fz1 = (std::max) (fz1, q.z + hz1);
							fx1 = (std::max) (fx1, std::fabs (q.x) + hx);
							}
				}
			// THE STOCK BOUNDARY the stock setup draws (lstock_id), mapped onto
			// (z, radius) whichever way best covers where the tool feeds.
			double rawR = 0;
			for (const OpPath &p : paths)
				for (const Seg &s : p.segs)
					if (s.feed)
						rawR = (std::max) ({ rawR, std::fabs (s.a.x), std::fabs (s.b.x) });
			std::vector<P> outline;
			bool swap = false, dia = false;
			for (int k = 0; k < 2; ++k)
				{
				if (s3.lstock_id[k] <= 0)
					continue;
				const Bdry bd = ReadBoundary (s3.lstock_id[k]);
				LogIf (part, L"stock sim group " + std::to_wstring (g->grp_idn) + L" lstock_id[" + std::to_wstring (k)
									 + L"]: " + bd.info);
				if (!bd.ok || !outline.empty ())
					continue;
				double best = -1;
				for (int m = 0; m < 4; ++m)
					{
					const std::vector<P> c = Mapped (bd.raw, (m & 1) != 0, (m & 2) != 0);
					double cz0 = 1e300, cz1 = -1e300, cr = 0;
					for (const P &q : c)
						{
						cz0 = (std::min) (cz0, q.z);
						cz1 = (std::max) (cz1, q.z);
						cr = (std::max) (cr, q.x);
						}
					// How much of the feed moves' Z span it covers, and how near its OD
					// is to the furthest the tool feeds out.
					const double zo = fz1 > fz0 ? (std::max) (0.0, (std::min) (cz1, fz1) - (std::max) (cz0, fz0)) / (fz1 - fz0) : 0;
					const double ro = cr > 0 && rawR > 0 ? (std::min) (cr, rawR) / (std::max) (cr, rawR) : 0;
					const double score = zo + ro;
					LogIf (part, L"stock sim   mapping " + std::wstring (m & 1 ? L"y->Z" : L"x->Z")
										 + (m & 2 ? L" diameter" : L" radius") + L": Z " + F (cz0) + L".." + F (cz1)
										 + L" R to " + F (cr) + L" = " + F (RevolvedVolume (c), 3) + L" in^3, score "
										 + F (score, 3));
					if (score > best + 1e-9)
						{
						best = score;
						swap = (m & 1) != 0;
						dia = (m & 2) != 0;
						outline = c;
						}
					}
				}

			std::wstring from;
			if (!outline.empty ())
				{
				zA = 1e300;
				zB = -1e300;
				for (const P &q : outline)
					{
					zA = (std::min) (zA, q.z);
					zB = (std::max) (zB, q.z);
					odR = (std::max) (odR, q.x);
					}
				from = std::wstring (L"stock boundary (") + (swap ? L"y->Z" : L"x->Z") + (dia ? L", diameter)" : L", radius)");
				}
			else if (bar != nullptr)
				{
				odR = bar->od / 2;
				idR = bar->holeInStock ? bar->id / 2 : 0;
				zA = bar->refAtMaxZ ? bar->refZ - bar->length : bar->refZ;
				zB = zA + bar->length;
				from = bar == &s3.stock_def[0] ? L"left bar" : L"right bar";
				}
			else
				{
				odR = fx1;
				zA = fz0;
				zB = fz1;
				from = L"NO BAR DEFINED - feed-move extents (fallback)";
				}
			// The raster: the stock and every feed move, a little beyond.
			const double zLo = (std::min) (zA, fz0) - 0.1, zHi = (std::max) (zB, fz1) + 0.1;
			const double rHi = (std::max) (odR, fx1) + 0.1;
			double h = 0.005;
			while ((rHi / h) * ((zHi - zLo) / h) > 8e6)
				h *= 1.25;
			Grid ins, nose;
			ins.Init (h, zLo, zHi, rHi);
			nose.Init (h, zLo, zHi, rHi);
			if (!outline.empty ())
				{
				ins.FillPolygon (outline);
				nose.FillPolygon (outline);
				}
			else
				{
				ins.FillStock (idR, odR, zA, zB);
				nose.FillStock (idR, odR, zA, zB);
				}
			const double stockExact = !outline.empty () ? RevolvedVolume (outline)
														: kPi * (odR * odR - idR * idR) * (zB - zA);
			const double stockGrid = ins.Volume ();
			LogIf (part, L"stock sim group " + std::to_wstring (g->grp_idn) + L" stock from " + from + L": R "
								 + F (idR) + L".." + F (odR) + L" Z " + F (zA) + L".." + F (zB) + L" = " + F (stockExact, 3)
								 + L" in^3 (raster " + F (stockGrid, 3) + L") | raster " + std::to_wstring (ins.nx) + L"x"
								 + std::to_wstring (ins.nz) + L" cells of " + F (h) + L" | feed moves Z " + F (fz0) + L".."
								 + F (fz1) + L" R to " + F (fx1));

			double totIns = 0, totNose = 0, totAir = 0, totSwept = 0;	// air / all feed seconds
			std::set<long> mcLogged;
			double mcPrev = stockExact;		// Mastercam's stock left so far, by its boundaries
			for (const OpPath &p : paths)
				{
				const auto s0 = std::chrono::steady_clock::now ();
				const Shape &sh = shapes[p.op->tl.slot];
				ins.BeginOp ();
				nose.BeginOp ();
				double rapidIn = 0, cutT = 0, airT = 0, cutL = 0, airL = 0;
				long feeds = 0;
				for (const Seg &s : p.segs)
					{
					if (!s.feed)
						{
						rapidIn += ins.ThroughStock (s.a, s.b);
						continue;
						}
					// AIR CUTTING BY TIME: the move in pieces of at most 0.05", each timed
					// like the path walk and counted as cutting when it clears more than
					// a sliver (a mean depth over 0.001", and more than two cells).
					const double len = std::hypot (s.b.z - s.a.z, s.b.x - s.a.x);
					const int k = (std::max) (1, static_cast<int> (std::ceil (len / 0.05)));
					for (int q = 0; q < k; ++q)
						{
						const P pa { s.a.z + (s.b.z - s.a.z) * q / k, s.a.x + (s.b.x - s.a.x) * q / k };
						const P pb { s.a.z + (s.b.z - s.a.z) * (q + 1) / k, s.a.x + (s.b.x - s.a.x) * (q + 1) / k };
						const double before = ins.removedArea;
						ins.Sweep (sh.insert, pa, pb);
						nose.Sweep (sh.nose, pa, pb);
						++feeds;
						const double cut = ins.removedArea - before;
						const double t = Paths::FeedSeconds (s.speed, s.css, s.cap, p.op->tl.mm, len / k,
															 std::fabs ((pa.x + pb.x) / 2), s.rate);
						const bool cutting = cut > (std::max) (2.5 * h * h, 0.001 * len / k);
						(cutting ? cutT : airT) += t;
						(cutting ? cutL : airL) += len / k;
						}
					}
				const double secs = std::chrono::duration<double> (std::chrono::steady_clock::now () - s0).count ();
				totIns += ins.removedVol;
				totNose += nose.removedVol;
				totAir += airT;
				totSwept += airT + cutT;
				// Feed time, for an average removal rate to set beside 12*SFM*IPR*DOC.
				const Paths::Totals pt = Paths::Walk (*p.op);
				const double minutes = pt.feedSeconds / 60.0;
				const double swept = ins.airArea + ins.removedArea;
				// The op's own stock boundary, as Mastercam keeps it (lathe ops).
				std::wstring mc;
				const long ol = p.op->cmn_lathe.lstock_id;
				if (ol > 0 || p.op->cmn_lathe.rstock_id > 0)
					{
					mc = L" | Mastercam bdry L" + std::to_wstring (ol) + L" R" + std::to_wstring (p.op->cmn_lathe.rstock_id)
						 + L" valid " + std::to_wstring (p.op->cmn_lathe.bdrys_valid) + L" upd "
						 + std::to_wstring (p.op->cmn_lathe.upd_cur_bdry) + L"/" + std::to_wstring (p.op->cmn_lathe.upd_subs_bdry);
					const Bdry ob = ReadBoundary (ol);
					if (ob.ok)
						{
						const double v = RevolvedVolume (Mapped (ob.raw, swap, dia));
						mc += L": " + F (v, 3) + L" in^3 left, removed " + F (mcPrev - v, 3) + L" in^3";
						mcPrev = v;
						}
					if (!mcLogged.count (ol))
						LogIf (part, L"stock sim   op " + std::to_wstring (p.op->op_idn) + L" bdry " + ob.info);
					mcLogged.insert (ol);
					}
				Result &res = results[p.op->op_idn];
				res.removed = ins.removedVol;
				res.airPct = airT + cutT > 0 ? 100.0 * airT / (airT + cutT) : 0;
				res.ok = p.holes == 0 && p.otherType == 0 && !p.afterFlip;
				LogIf (part, L"stock sim op " + std::to_wstring (p.op->op_idn) + L" T"
									 + std::to_wstring (p.op->tl.tlno) + L" \""
									 + std::wstring (p.op->comment, wcsnlen (p.op->comment, COMMENT_SIZE))
									 + L"\": removed " + F (ins.removedVol, 3) + L" in^3 (nose only "
									 + F (nose.removedVol, 3) + L"), AIR " + F (airT + cutT > 0 ? 100.0 * airT / (airT + cutT) : 0, 1)
									 + L"% of feed time (" + Ms (airT) + L" of " + Ms (airT + cutT) + L"), "
									 + F (airL + cutL > 0 ? 100.0 * airL / (airL + cutL) : 0, 1) + L"% of feed length ("
									 + F (airL, 2) + L" of " + F (airL + cutL, 2) + L" in) | swept " + F (swept, 3)
									 + L" in^2, cells " + std::to_wstring (
										   static_cast<long long> (ins.removedArea / (h * h) + 0.5))
									 + L" | feed " + Ms (pt.feedSeconds) + L", avg " + F (minutes > 0 ? ins.removedVol / minutes : 0, 3)
									 + L" in^3/min | " + std::to_wstring (feeds) + L" feed pieces, rapids through stock "
									 + F (rapidIn, 3) + L" in" + (p.holes ? L", " + std::to_wstring (p.holes) + L" DRILL HOLES NOT SIMULATED" : L"")
									 + (p.otherType ? L", " + std::to_wstring (p.otherType) + L" non-lathe moves skipped" : L"")
									 + L" | " + F (secs * 1000, 0) + L" ms" + mc);
				}
			LogIf (part, L"stock sim group " + std::to_wstring (g->grp_idn) + L" TOTAL removed " + F (totIns, 3)
								 + L" in^3 (nose only " + F (totNose, 3) + L"), left " + F (ins.Volume (), 3)
								 + L" in^3 of " + F (stockGrid, 3) + L" (nose only left " + F (nose.Volume (), 3)
								 + L"), AIR " + F (totSwept > 0 ? 100.0 * totAir / totSwept : 0, 1) + L"% of feed time ("
								 + Ms (totAir) + L" of " + Ms (totSwept) + L")");
			}
		LogIf (part, L"stock sim done in "
							 + F (std::chrono::duration<double> (std::chrono::steady_clock::now () - t0).count (), 2) + L" s");
		return results;
		}
	}
