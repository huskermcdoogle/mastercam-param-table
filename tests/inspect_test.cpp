// Tool inspection: m:ss in and out, and why each stop happened.
// Synthetic toolpaths shaped like a real rough: a 9:00 timer stopping only between
// cuts (so each stop lands a little past 9:00), plus the stop at the end.
#include "../src/Inspect.h"

#include <cstdio>
#include <cmath>

static int failed = 0;
static void Check (bool ok, const char *what)
	{
	std::printf ("  %s  %s\n", ok ? "ok  " : "FAIL", what);
	if (!ok)
		++failed;
	}

static Paths::Totals Path (double cut, double seconds, std::initializer_list<std::pair<double, double>> stops,
						   const wchar_t *comment = L"CHANGE/ROTATE INSERT")
	{
	Paths::Totals t;
	t.ok = true;
	t.cutLength = cut;
	t.feedSeconds = seconds;
	for (const auto &s : stops)
		{
		Paths::Totals::Inspection in;
		in.cutLength = s.first;
		in.feedSeconds = s.second;
		in.comment = comment;
		t.inspections.push_back (in);
		}
	return t;
	}

int main ()
	{
	double v = 0;
	Check (Inspect::MinSec (540) == L"9:00", "540 s shows as 9:00");
	Check (Inspect::MinSec (3723) == L"1:02:03", "3723 s shows as 1:02:03");
	Check (Inspect::MinSec (90.5) == L"1:30.5", "90.5 s keeps its tenth");
	Check (Inspect::ParseMinSec (L"9:00", v) && v == 540, "9:00 reads as 540 s");
	Check (Inspect::ParseMinSec (L" 1:02:30 ", v) && v == 3750, "1:02:30 reads as 3750 s");
	Check (Inspect::ParseMinSec (L"540", v) && v == 540, "540 reads as seconds");
	Check (!Inspect::ParseMinSec (L"9:75", v), "9:75 is refused");
	Check (!Inspect::ParseMinSec (L"nine", v), "text is refused");
	Check (Inspect::IsFlip (L"CHANGE/ROTATE INSERT") && Inspect::IsFlip (L"flip insert"), "flip comments");
	Check (!Inspect::IsFlip (L"CHECK SIZE"), "an inspection that is not a flip");

	// A 9:00 timer, between cuts, stop at end: three timed stops then the end one.
	Inspect::Settings s;
	s.doStop = s.timeOn = s.atEnd = s.betweenCuts = true;
	s.time = 540;
	Paths::Totals t = Path (40.0, 1900, { { 10.9, 558 }, { 21.6, 1118 }, { 32.4, 1690 }, { 40.0, 1900 } });
	Inspect::Result r = Inspect::Explain (s, t);
	Check (r.stops == 4 && r.flips == 4, "four stops, all flips");
	Check (r.byTime == 3 && r.atEnd == 1 && r.other == 0, "three by time, one at the end");
	Check (r.why == L"time 3 + end 1", "why: time 3 + end 1");
	Check (std::fabs (r.longest - 572) < 1e-9, "longest between flips is 9:32");
	Check (r.mode == L"between cuts", "mode between cuts");

	// No end stop wanted, and a stop well short of the timer: not time, so cuts.
	s.atEnd = false;
	t = Path (12.0, 700, { { 3.0, 200 }, { 11.0, 650 } });
	r = Inspect::Explain (s, t);
	Check (r.byTime == 0 && r.other == 2, "stops short of the timer count as cuts");

	// Inspections that are not flips are counted as stops only.
	t = Path (12.0, 700, { { 6.0, 560 } }, L"CHECK SIZE");
	r = Inspect::Explain (s, t);
	Check (r.stops == 1 && r.flips == 0, "an inspection stop is not a flip");

	// Inspection on, nothing in the toolpath.
	t = Path (5.0, 300, {});
	r = Inspect::Explain (s, t);
	Check (r.flips == 0 && r.why == L"inspection on, no stop in the toolpath", "on but no stops - said so");

	// Distance: every 10 of cut.
	Inspect::Settings d;
	d.doStop = d.distOn = true;
	d.dist = 10;
	d.betweenCuts = false;
	d.minCut = 0.25;
	t = Path (25.0, 900, { { 10.2, 360 }, { 20.4, 720 } });
	r = Inspect::Explain (d, t);
	Check (r.byDist == 2, "two by distance");
	Check (r.mode == L"mid-cut, finishes the pass if under 0.25 left", "mode mid-cut with min cut");

	if (failed)
		std::printf ("inspect_test: %d FAILED\n", failed);
	else
		std::printf ("inspect_test: all checks passed\n");
	return failed ? 1 : 0;
	}
