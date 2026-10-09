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
#include <vector>

namespace
	{
	using namespace StockRaster;

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
		Util::Log (part, line);
		return s;
		}

	/// One straight piece of the path (arcs come as chords), control point.
	struct Seg
		{
		P a, b;
		bool feed;
		};

	struct OpPath
		{
		operation *op = nullptr;
		std::vector<Seg> segs;
		long holes = 0;			//!< canned drill holes - not simulated
		long otherType = 0;		//!< moves that are not lathe NCI (mill moves in an MT group)
		};

	OpPath Read (operation &op)
		{
		OpPath r;
		r.op = &op;
		CBnciReadWrite nci;
		const INT_PTR n = nci.ReadSection (op.op_idn, true);
		bool started = false;
		P at { 0, 0 };
		for (INT_PTR i = 0; i < n; ++i)
			{
			const nci_bin *b = nci[i];
			if (b == nullptr)
				continue;
			const NCI_GCODE g = b->gcode;
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
				r.segs.push_back ({ at, end, g == NCI_LINEAR });
			else
				{
				P prev = at;
				for (const P &p : Chords (at, end, { b->u.l2.cpt[1], b->u.l2.cpt[0] }, g == NCI_ARC_CW, 0.0005))
					{
					r.segs.push_back ({ prev, p, true });
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
	void Run (const std::filesystem::path &part)
		{
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
				Util::Log (part, head + L"(none found) - " + std::to_wstring (grp.second.size ()) + L" ops skipped");
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
			Util::Log (part, head);
			if (g->product != PRODUCT_LATHE && g->product != PRODUCT_MT)
				continue;

			// The paths, and the tools' shapes.
			std::vector<OpPath> paths;
			for (operation *op : grp.second)
				{
				const long code = static_cast<long> (op->opcode);
				if (code == TP_LSTOCK_FLIP || code == TP_LSTOCK_XFER)
					Util::Log (part, L"stock sim op " + std::to_wstring (op->op_idn) + L": STOCK "
									 + (code == TP_LSTOCK_FLIP ? L"FLIP" : L"TRANSFER")
									 + L" - not followed; what comes after is in the wrong frame");
				if (op->db.nci_flag)
					{
					Util::Log (part, L"stock sim op " + std::to_wstring (op->op_idn) + L": needs regen - skipped");
					continue;
					}
				OpPath p = Read (*op);
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
			std::wstring from;
			if (bar != nullptr)
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
			ins.FillStock (idR, odR, zA, zB);
			nose.FillStock (idR, odR, zA, zB);
			const double stockExact = kPi * (odR * odR - idR * idR) * (zB - zA);
			const double stockGrid = ins.Volume ();
			Util::Log (part, L"stock sim group " + std::to_wstring (g->grp_idn) + L" stock from " + from + L": R "
								 + F (idR) + L".." + F (odR) + L" Z " + F (zA) + L".." + F (zB) + L" = " + F (stockExact, 3)
								 + L" in^3 (raster " + F (stockGrid, 3) + L") | raster " + std::to_wstring (ins.nx) + L"x"
								 + std::to_wstring (ins.nz) + L" cells of " + F (h) + L" | feed moves Z " + F (fz0) + L".."
								 + F (fz1) + L" R to " + F (fx1));

			double totIns = 0, totNose = 0, totAir = 0, totSwept = 0;
			for (const OpPath &p : paths)
				{
				const auto s0 = std::chrono::steady_clock::now ();
				const Shape &sh = shapes[p.op->tl.slot];
				ins.BeginOp ();
				nose.BeginOp ();
				double rapidIn = 0;
				long feeds = 0;
				for (const Seg &s : p.segs)
					{
					if (!s.feed)
						{
						rapidIn += ins.ThroughStock (s.a, s.b);
						continue;
						}
					++feeds;
					ins.Sweep (sh.insert, s.a, s.b);
					nose.Sweep (sh.nose, s.a, s.b);
					}
				const double secs = std::chrono::duration<double> (std::chrono::steady_clock::now () - s0).count ();
				totIns += ins.removedVol;
				totNose += nose.removedVol;
				totAir += ins.airArea;
				totSwept += ins.airArea + ins.removedArea;
				// Feed time, for an average removal rate to set beside 12*SFM*IPR*DOC.
				const Paths::Totals pt = Paths::Walk (*p.op);
				const double minutes = pt.feedSeconds / 60.0;
				const double swept = ins.airArea + ins.removedArea;
				Util::Log (part, L"stock sim op " + std::to_wstring (p.op->op_idn) + L" T"
									 + std::to_wstring (p.op->tl.tlno) + L" \""
									 + std::wstring (p.op->comment, wcsnlen (p.op->comment, COMMENT_SIZE))
									 + L"\": removed " + F (ins.removedVol, 3) + L" in^3 (nose only "
									 + F (nose.removedVol, 3) + L"), air " + F (swept > 0 ? 100.0 * ins.airArea / swept : 0, 1)
									 + L"% of " + F (swept, 3) + L" in^2 swept, cells " + std::to_wstring (
										   static_cast<long long> (ins.removedArea / (h * h) + 0.5))
									 + L" | feed " + Ms (pt.feedSeconds) + L", avg " + F (minutes > 0 ? ins.removedVol / minutes : 0, 3)
									 + L" in^3/min | " + std::to_wstring (feeds) + L" feed pieces, rapids through stock "
									 + F (rapidIn, 3) + L" in" + (p.holes ? L", " + std::to_wstring (p.holes) + L" DRILL HOLES NOT SIMULATED" : L"")
									 + (p.otherType ? L", " + std::to_wstring (p.otherType) + L" non-lathe moves skipped" : L"")
									 + L" | " + F (secs * 1000, 0) + L" ms");
				}
			Util::Log (part, L"stock sim group " + std::to_wstring (g->grp_idn) + L" TOTAL removed " + F (totIns, 3)
								 + L" in^3 (nose only " + F (totNose, 3) + L"), left " + F (ins.Volume (), 3)
								 + L" in^3 of " + F (stockGrid, 3) + L" (nose only left " + F (nose.Volume (), 3)
								 + L"), air " + F (totSwept > 0 ? 100.0 * totAir / totSwept : 0, 1) + L"%");
			}
		Util::Log (part, L"stock sim done in "
							 + F (std::chrono::duration<double> (std::chrono::steady_clock::now () - t0).count (), 2) + L" s");
		}
	}
