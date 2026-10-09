// The stock raster (StockRaster.h) against volumes worked out by hand.
#include "../src/StockRaster.h"

#include <cmath>
#include <cstdio>

int main ()
	{
	using namespace StockRaster;
	int failed = 0;
	auto near = [&failed] (double got, double want, double tol, const char *what)
		{
		const bool ok = std::fabs (got - want) <= tol;
		std::printf ("%s  %s: %.5f (want %.5f)\n", ok ? "ok  " : "FAIL", what, got, want);
		if (!ok)
			++failed;
		};

	// A bar R 1, Z 0..2.
	Grid g;
	g.Init (0.005, -0.5, 2.5, 1.3);
	g.FillStock (0, 1, 0, 2);
	near (g.Volume (), kPi * 2, 0.01, "stock volume");

	// A square 0.2 insert, the control point at its lower left corner.
	const std::vector<P> sq = { { 0, 0 }, { 0.2, 0 }, { 0.2, 0.2 }, { 0, 0.2 } };

	// Turn to R 0.9 the whole length: pi (1 - 0.81) 2.
	g.BeginOp ();
	g.Sweep (sq, { 2.1, 0.9 }, { -0.1, 0.9 });
	near (g.removedVol, kPi * 0.19 * 2, 0.01, "turn pass removed");
	near (g.removedArea, 0.1 * 2, 0.005, "turn pass area");
	near (g.airArea + g.removedArea, 2.4 * 0.2, 0.01, "turn pass swept area (stadium of a square = rectangle)");

	// The same pass again cuts only air.
	g.BeginOp ();
	g.Sweep (sq, { 2.1, 0.9 }, { -0.1, 0.9 });
	near (g.removedVol, 0, 1e-9, "second pass removed");

	// Face 0.1 off the end, past centre: the disc R 0.9 x 0.1.
	g.BeginOp ();
	g.Sweep (sq, { 1.9, 1.2 }, { 1.9, -0.1 });
	near (g.removedVol, kPi * 0.81 * 0.1, 0.005, "face pass removed");
	near (g.Volume (), kPi * 2 - kPi * 0.19 * 2 - kPi * 0.81 * 0.1, 0.01, "left after both");

	// A diagonal move: the hull of the square at both ends is a hexagon -
	// square + parallelogram 0.2 x 0.5 on each side = 0.04 + 2 * 0.5 * 0.2 / ... (shoelace).
	{
	std::vector<P> pts;
	for (const P &p : sq)
		pts.push_back (p);
	for (const P &p : sq)
		pts.push_back ({ p.z + 0.5, p.x + 0.5 });
	const std::vector<P> h = Hull (pts);
	double a = 0;
	for (size_t i = 0; i < h.size (); ++i)
		{
		const P &p = h[i], &q = h[(i + 1) % h.size ()];
		a += p.z * q.x - q.z * p.x;
		}
	near (static_cast<double> (h.size ()), 6, 0, "diagonal hull corners");
	near (a / 2, 0.04 + 2 * 0.5 * 0.2, 1e-12, "diagonal hull area");
	}

	// A quarter arc R 1 about (0, 0), CCW from +Z to +X: chords stay on it and end where told.
	{
	const std::vector<P> c = Chords ({ 1, 0 }, { 0, 1 }, { 0, 0 }, false, 0.0005);
	double worst = 0;
	for (const P &p : c)
		worst = (std::max) (worst, std::fabs (std::hypot (p.z, p.x) - 1));
	near (worst, 0, 1e-9, "arc chord points on the arc");
	near (c.back ().z + c.back ().x, 1, 1e-12, "arc ends at its end");
	// The same arc CW goes the long way round: 3/4 of a turn.
	const std::vector<P> w = Chords ({ 1, 0 }, { 0, 1 }, { 0, 0 }, true, 0.0005);
	near (static_cast<double> (w.size ()) / c.size (), 3, 0.05, "CW arc is the long way (3x the pieces)");
	}

	// A cell swept twice in one op counts once.
	{
	Grid t;
	t.Init (0.01, 0, 1, 1);
	t.FillStock (0, 0.5, 0, 1);
	t.BeginOp ();
	t.Sweep (sq, { 0.2, 0.6 }, { 0.2, 0.6 });
	t.Sweep (sq, { 0.2, 0.6 }, { 0.2, 0.6 });
	near (t.airArea, 0.04, 0.002, "air counted once per op");
	}

	if (failed)
		std::printf ("%d FAILED\n", failed);
	return failed ? 1 : 0;
	}
