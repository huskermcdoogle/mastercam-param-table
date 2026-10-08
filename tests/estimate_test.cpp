// The live estimate formula against the C++ model it is built from.
//
// Writes estimate.xlsx (one operation: feed, feed_mode, speed, speed_mode,
// max_ss and the est formula) and estimate_expect.txt: per scenario, the cell
// edits and the seconds the model expects. tools/check_estimate.ps1 makes the
// same edits in real Excel and compares.
#include "../src/Estimate.h"
#include "../src/Xlsx.h"
#include "../src/Csv.h"

#include <cmath>
#include <cstdio>

int main (int argc, char **argv)
	{
	using namespace Estimate;
	const std::string dir = argc > 1 ? argv[1] : ".";
	int failed = 0;
	auto check = [&failed] (bool ok, const char *what)
		{
		if (!ok)
			{
			std::printf ("FAIL  %s\n", what);
			++failed;
			}
		};

	// ---- Bands keep total length and length x diameter.
	std::vector<std::pair<double, double>> pieces;
	for (int i = 0; i < 200; ++i)
		pieces.push_back ({ 0.5 + 0.01 * i, 1.0 + 0.03 * i });	// (length, radius)
	double L = 0, LD = 0;
	for (const auto &p : pieces)
		{
		L += p.first;
		LD += p.first * 2.0 * p.second;
		}
	const std::vector<Band> bands = Bands (pieces, 8);
	double bl = 0, bld = 0;
	for (const Band &b : bands)
		{
		bl += b.len;
		bld += b.lenDia;
		}
	check (bands.size () <= 8 && bands.size () >= 6, "bands: at most the number asked for");
	check (std::fabs (bl - L) < 1e-9 && std::fabs (bld - LD) < 1e-9, "bands keep length and length x diameter");

	// ---- The group behind the sample sheet: per rev, CSS, a cap that some
	// bands reach (small diameters) and some do not.
	Group g;
	g.feed = -0.01;
	g.speed = 200;
	g.css = true;
	g.cap = 500;
	g.bands = bands;

	Cells c;
	c.feed = L"B3";
	c.feedMode = L"C3";
	c.speed = L"D3";
	c.css = L"E3";
	c.cssIsWord = true;
	c.cap = L"F3";

	// A second, fixed group (not on the sheet): it must not move.
	Group fixed;
	fixed.feed = 20;			// per minute
	fixed.bands = Bands ({ { 10.0, 3.0 } }, 4);
	const double fixedSec = Seconds (fixed);

	const double base = Seconds (g) + fixedSec;
	const std::wstring formula = Term (g, c) + L"+" + Term (fixed, Cells ());

	Xlsx::Sheet s;
	s.rows = { { L"op_idn", L"feed", L"feed_mode", L"speed", L"speed_mode", L"max_ss", L"est" },
			   { L"1", L"0.01", L"per rev", L"200", L"CSS", L"500", Csv::Tidy (base) } };
	s.text = { 0, 0, 1, 0, 1, 0, 0 };
	s.formula = { { L"", L"", L"", L"", L"", L"", formula } };
	s.groupNames = { L"x" };
	s.group.assign (7, 0);
	if (!Xlsx::Write (dir + "/estimate.xlsx", s))
		{
		std::puts ("FAIL: write");
		return 1;
		}

	// ---- Scenarios: edits, and what the model says.
	FILE *out = nullptr;
	fopen_s (&out, (dir + "/estimate_expect.txt").c_str (), "w");
	auto scenario = [&] (const char *edits, Group gg)
		{
		std::fprintf (out, "%s|%.6f\n", edits, Seconds (gg) + fixedSec);
		};
	scenario ("", g);
	{ Group x = g; x.feed = -0.02;                scenario ("B3=0.02", x); }
	{ Group x = g; x.speed = 400;                 scenario ("D3=400", x); }
	{ Group x = g; x.css = false;                 scenario ("E3=RPM", x); }
	{ Group x = g; x.cap = 50;                    scenario ("F3=50", x); }
	{ Group x = g; x.cap = 0;                     scenario ("F3=0", x); }
	{ Group x = g; x.feed = 5;                    scenario ("C3=per min;B3=5", x); }
	{ Group x = g; x.feed = -0.015; x.speed = 300; x.css = false; scenario ("B3=0.015;D3=300;E3=RPM", x); }
	std::fclose (out);

	check (std::fabs (base - (Seconds (g) + fixedSec)) < 1e-9, "base is the dumped model");
	if (failed)
		{
		std::printf ("estimate_test: %d check(s) FAILED\n", failed);
		return 1;
		}
	std::printf ("estimate_test ok (base %.1f s)\n", base);
	return 0;
	}
