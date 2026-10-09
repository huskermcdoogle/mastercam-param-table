//
// StockRaster.h - the 2D lathe stock raster behind StockSim (no SDK, so the
// unit tests can drive it): cells of the XZ half-plane, rows by radius, each a
// ring of 2*pi*r*h^2 (Pappus). A tool's convex outline swept along a straight
// move is the hull of the outline at both ends.
//
#pragma once

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <vector>

namespace StockRaster
	{
	const double kPi = 3.14159265358979323846;

	/// A point in the lathe half-plane: z along the spindle, x a RADIUS.
	struct P
		{
		double z, x;
		};

	/// Convex hull (monotone chain), anticlockwise, no repeated points.
	inline std::vector<P> Hull (std::vector<P> pts)
		{
		std::sort (pts.begin (), pts.end (), [] (const P &a, const P &b)
				   { return a.z < b.z || (a.z == b.z && a.x < b.x); });
		if (pts.size () < 3)
			return pts;
		auto cross = [] (const P &o, const P &a, const P &b)
			{ return (a.z - o.z) * (b.x - o.x) - (a.x - o.x) * (b.z - o.z); };
		std::vector<P> h (2 * pts.size ());
		size_t k = 0;
		for (size_t i = 0; i < pts.size (); ++i)
			{
			while (k >= 2 && cross (h[k - 2], h[k - 1], pts[i]) <= 0)
				--k;
			h[k++] = pts[i];
			}
		for (size_t i = pts.size () - 1, t = k + 1; i-- > 0;)
			{
			while (k >= t && cross (h[k - 2], h[k - 1], pts[i]) <= 0)
				--k;
			h[k++] = pts[i];
			}
		h.resize (k - 1);
		return h;
		}

	/// How far the point is from the convex polygon's outline (negative inside).
	inline double SignedDistance (const std::vector<P> &poly, P p)
		{
		if (poly.empty ())
			return 0;
		double best = 1e300;
		bool inside = poly.size () >= 3;
		for (size_t i = 0; i < poly.size (); ++i)
			{
			const P &a = poly[i], &b = poly[(i + 1) % poly.size ()];
			const double ez = b.z - a.z, ex = b.x - a.x;
			const double len2 = ez * ez + ex * ex;
			double t = len2 > 0 ? ((p.z - a.z) * ez + (p.x - a.x) * ex) / len2 : 0;
			t = (std::max) (0.0, (std::min) (1.0, t));
			const double dz = a.z + t * ez - p.z, dx = a.x + t * ex - p.x;
			best = (std::min) (best, std::sqrt (dz * dz + dx * dx));
			if (ez * (p.x - a.x) - ex * (p.z - a.z) < 0)
				inside = false;
			}
		return inside ? -best : best;
		}

	/// THE STOCK: a raster of the half-plane, rows by radius from the axis.
	struct Grid
		{
		double h = 0.005, z0 = 0;
		int nx = 0, nz = 0;
		std::vector<uint8_t> mat;			//!< 1 = material still there
		std::vector<uint32_t> stamp;		//!< last op that swept the cell
		uint32_t opStamp = 0;

		// Per op.
		double removedVol = 0, removedArea = 0, airArea = 0;

		void Init (double h_, double zLo, double zHi, double rHi)
			{
			h = h_;
			z0 = zLo;
			nz = static_cast<int> (std::ceil ((zHi - zLo) / h));
			nx = static_cast<int> (std::ceil (rHi / h));
			mat.assign (static_cast<size_t> (nx) * nz, 0);
			stamp.assign (mat.size (), 0);
			}

		void FillStock (double idR, double odR, double zA, double zB)
			{
			for (int i = 0; i < nx; ++i)
				{
				const double r = (i + 0.5) * h;
				if (r < idR || r > odR)
					continue;
				for (int j = 0; j < nz; ++j)
					{
					const double z = z0 + (j + 0.5) * h;
					if (z >= zA && z <= zB)
						mat[static_cast<size_t> (i) * nz + j] = 1;
					}
				}
			}

		double Volume () const
			{
			double v = 0;
			for (int i = 0; i < nx; ++i)
				{
				long long c = 0;
				for (int j = 0; j < nz; ++j)
					c += mat[static_cast<size_t> (i) * nz + j];
				v += c * 2.0 * kPi * (i + 0.5) * h * h * h;
				}
			return v;
			}

		void BeginOp ()
			{
			++opStamp;
			removedVol = removedArea = airArea = 0;
			}

		/// Clear every cell whose centre lies in the convex polygon - folded
		/// about the axis, since a cell is a ring.
		void Clear (const std::vector<P> &poly)
			{
			double xLo = 1e300, xHi = -1e300;
			for (const P &p : poly)
				{
				xLo = (std::min) (xLo, p.x);
				xHi = (std::max) (xHi, p.x);
				}
			const int kLo = static_cast<int> (std::floor (xLo / h - 0.5)) - 1;
			const int kHi = static_cast<int> (std::ceil (xHi / h - 0.5)) + 1;
			for (int k = kLo; k <= kHi; ++k)
				{
				const double yc = (k + 0.5) * h;
				if (yc < xLo || yc > xHi)
					continue;
				const int row = k >= 0 ? k : -k - 1;
				if (row >= nx)
					continue;
				double zl = 1e300, zr = -1e300;
				for (size_t e = 0; e < poly.size (); ++e)
					{
					const P &a = poly[e], &b = poly[(e + 1) % poly.size ()];
					if ((a.x - yc) * (b.x - yc) > 0 || a.x == b.x)
						{
						if (a.x == yc)
							{
							zl = (std::min) (zl, a.z);
							zr = (std::max) (zr, a.z);
							}
						continue;
						}
					const double z = a.z + (yc - a.x) / (b.x - a.x) * (b.z - a.z);
					zl = (std::min) (zl, z);
					zr = (std::max) (zr, z);
					}
				if (zl > zr)
					continue;
				int jl = static_cast<int> (std::ceil ((zl - z0) / h - 0.5));
				int jr = static_cast<int> (std::floor ((zr - z0) / h - 0.5));
				jl = (std::max) (jl, 0);
				jr = (std::min) (jr, nz - 1);
				const double ring = 2.0 * kPi * (row + 0.5) * h * h * h;
				uint8_t *m = &mat[static_cast<size_t> (row) * nz];
				uint32_t *s = &stamp[static_cast<size_t> (row) * nz];
				for (int j = jl; j <= jr; ++j)
					{
					if (s[j] == opStamp)
						continue;
					s[j] = opStamp;
					if (m[j])
						{
						m[j] = 0;
						removedVol += ring;
						removedArea += h * h;
						}
					else
						airArea += h * h;
					}
				}
			}

		/// Sweep the tool's shape along a straight move.
		void Sweep (const std::vector<P> &shape, P a, P b)
			{
			std::vector<P> pts;
			pts.reserve (shape.size () * 2);
			for (const P &p : shape)
				pts.push_back ({ p.z + a.z, p.x + a.x });
			if (a.z != b.z || a.x != b.x)
				for (const P &p : shape)
					pts.push_back ({ p.z + b.z, p.x + b.x });
			Clear (pts.size () == shape.size () ? pts : Hull (pts));
			}

		/// How far a rapid's control point travels through material.
		double ThroughStock (P a, P b) const
			{
			const double len = std::hypot (b.z - a.z, b.x - a.x);
			const int n = (std::max) (1, static_cast<int> (std::ceil (len / (h * 0.5))));
			int hits = 0;
			for (int k = 0; k <= n; ++k)
				{
				const double z = a.z + (b.z - a.z) * k / n, x = std::fabs (a.x + (b.x - a.x) * k / n);
				const int i = static_cast<int> (x / h), j = static_cast<int> ((z - z0) / h);
				if (i >= 0 && i < nx && j >= 0 && j < nz && mat[static_cast<size_t> (i) * nz + j])
					++hits;
				}
			return len * hits / (n + 1);
			}
		};

	/// A lathe arc (G18: swept from Z towards X, anticlockwise in (z, x) for
	/// CCW) as chords within `tol` of it - the points after `from`, ending at `to`.
	inline std::vector<P> Chords (P from, P to, P c, bool cw, double tol)
		{
		const double su = from.z - c.z, sv = from.x - c.x, eu = to.z - c.z, ev = to.x - c.x;
		const double rad = std::sqrt (su * su + sv * sv);
		double sweep = std::atan2 (ev, eu) - std::atan2 (sv, su);
		if (cw)
			sweep = -sweep;
		while (sweep <= 1e-12)
			sweep += 2.0 * kPi;
		const double step = rad > tol ? 2.0 * std::acos (1.0 - tol / rad) : kPi / 8;
		const int pieces = (std::min) (4000, (std::max) (1, static_cast<int> (std::ceil (sweep / step))));
		const double a0 = std::atan2 (sv, su), dir = cw ? -1.0 : 1.0;
		std::vector<P> out;
		for (int k = 1; k <= pieces; ++k)
			{
			const double a = a0 + dir * sweep * k / pieces;
			out.push_back (k == pieces ? to : P { c.z + rad * std::cos (a), c.x + rad * std::sin (a) });
			}
		return out;
		}
	}
