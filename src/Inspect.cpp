#include "Inspect.h"
#include "Csv.h"

#include <cmath>
#include <cwctype>

namespace Inspect
	{
	std::wstring MinSec (double seconds)
		{
		if (!(seconds >= 0))
			return std::wstring ();
		const double tenths = std::round (seconds * 10.0);
		const long long whole = static_cast<long long> (tenths) / 10;
		const int frac = static_cast<int> (static_cast<long long> (tenths) % 10);
		wchar_t buf[48];
		if (whole >= 3600)
			swprintf_s (buf, L"%lld:%02lld:%02lld", whole / 3600, (whole / 60) % 60, whole % 60);
		else
			swprintf_s (buf, L"%lld:%02lld", whole / 60, whole % 60);
		std::wstring s = buf;
		if (frac != 0)
			s += L"." + std::to_wstring (frac);
		return s;
		}

	bool ParseMinSec (const std::wstring &text, double &seconds)
		{
		std::wstring t;
		for (wchar_t c : text)
			if (!std::iswspace (c))
				t += c;
		if (t.empty ())
			return false;
		std::vector<std::wstring> parts (1);
		for (wchar_t c : t)
			if (c == L':')
				parts.emplace_back ();
			else
				parts.back () += c;
		if (parts.size () > 3)
			return false;
		double total = 0;
		for (size_t i = 0; i < parts.size (); ++i)
			{
			double v = 0;
			if (!Csv::ParseDouble (parts[i], v) || v < 0)
				return false;
			if (i > 0 && v >= 60)
				return false;				// "9:75" is not a time
			total = total * 60.0 + v;
			}
		seconds = total;
		return true;
		}

	bool CheckMinSec (const std::wstring &text, std::wstring &why)
		{
		double s = 0;
		if (ParseMinSec (text, s))
			return true;
		why = L"\"" + text + L"\" is not a time - type minutes:seconds (9:00) or seconds (540)";
		return false;
		}

	bool IsFlip (const std::wstring &comment)
		{
		std::wstring u;
		for (wchar_t c : comment)
			u += static_cast<wchar_t> (std::towupper (c));
		for (const wchar_t *w : { L"ROTATE", L"FLIP", L"CHANGE", L"INDEX" })
			if (u.find (w) != std::wstring::npos)
				return true;
		return false;
		}

	Result Explain (const Settings &s, const Paths::Totals &t)
		{
		Result r;
		r.mode = s.betweenCuts ? L"between cuts"
							   : L"mid-cut, finishes the pass if under " + Csv::Tidy (s.minCut) + L" left";
		double lastTime = 0, lastLen = 0, lastFlip = 0;
		for (size_t i = 0; i < t.inspections.size (); ++i)
			{
			const Paths::Totals::Inspection &in = t.inspections[i];
			++r.stops;
			const bool flip = IsFlip (in.comment);
			// The cause: the end stop first, then whichever threshold was crossed.
			// A little slack - the recorded positions are rounded.
			const bool end = s.atEnd && i + 1 == t.inspections.size ()
							 && std::fabs (in.cutLength - t.cutLength) <= 1e-3 * (std::max) (1.0, t.cutLength);
			// Mastercam's own clock for the timer runs a little ahead of our feed
			// time (stops up to ~10% early on a real part), so with no cut-count
			// trigger switched on, an early stop is still the timer's (or the
			// distance's, whichever is nearer its setting).
			const bool cutTriggers = s.cutsOn || s.firstCut || s.eachDepth || s.eachGroove || s.eachSection;
			const double rt = s.timeOn && s.time > 0 ? (in.feedSeconds - lastTime) / s.time : -1;
			const double rd = s.distOn && s.dist > 0 ? (in.cutLength - lastLen) / s.dist : -1;
			bool byTime = !end && rt >= 0.97;
			bool byDist = !end && !byTime && rd >= 0.97;
			if (!end && !byTime && !byDist && !cutTriggers && (rt >= 0 || rd >= 0))
				{
				byTime = rt >= rd;
				byDist = !byTime;
				}
			if (flip)
				{
				++r.flips;
				if (end) ++r.atEnd;
				else if (byTime) ++r.byTime;
				else if (byDist) ++r.byDist;
				else ++r.other;
				r.longest = (std::max) (r.longest, in.feedSeconds - lastFlip);
				lastFlip = in.feedSeconds;
				}
			lastTime = in.feedSeconds;
			lastLen = in.cutLength;
			}
		if (r.flips > 0)
			r.longest = (std::max) (r.longest, t.feedSeconds - lastFlip);
		r.tail = (std::max) (0.0, t.feedSeconds - lastFlip);
		auto add = [&r] (int n, const wchar_t *what)
			{
			if (n > 0)
				r.why += (r.why.empty () ? L"" : L" + ") + std::wstring (what) + L" " + std::to_wstring (n);
			};
		add (r.byTime, L"time");
		add (r.byDist, L"distance");
		add (r.other, L"cuts");
		add (r.atEnd, L"end");
		if (r.stops > r.flips)
			r.why += (r.why.empty () ? L"" : L"  ") + std::wstring (L"(")
					 + std::to_wstring (r.stops - r.flips) + L" stop(s) not a flip)";
		if (r.why.empty () && s.doStop)
			r.why = s.commentOn ? L"inspection on, no stop in the toolpath"
								: L"no comment on the stops - not counted";
		return r;
		}

	std::wstring Criteria (const Settings &s, const std::wstring &unit)
		{
		std::wstring o;
		auto add = [&o] (const std::wstring &t) { o += (o.empty () ? L"" : L", ") + t; };
		if (!s.doStop)
			return o;
		if (s.timeOn) add (L"every " + MinSec (s.time));
		if (s.distOn) add (L"every " + Csv::Tidy (s.dist) + (unit.empty () ? L"" : L" " + unit) + L" of cut");
		if (s.cutsOn) add (L"every " + std::to_wstring (s.cuts) + L" cut(s)");
		if (s.firstCut) add (L"after first cut");
		if (s.eachDepth) add (L"each depth");
		if (s.eachGroove) add (L"each groove");
		if (s.eachSection) add (L"every " + std::to_wstring (s.sections) + L" section(s)");
		if (s.atEnd) add (L"at end");
		return o;
		}

	std::wstring Describe (long opIdn, const Settings &s, const Result &r)
		{
		std::wstring o = L"inspect op " + std::to_wstring (opIdn) + L": " + (s.doStop ? L"on" : L"off");
		if (s.doStop && !Criteria (s, L"").empty ())
			o += L", " + Criteria (s, L"");
		o += L", " + r.mode + L" | " + std::to_wstring (r.stops) + L" stop(s), "
			 + std::to_wstring (r.flips) + L" flip(s)" + (r.why.empty () ? L"" : L": " + r.why)
			 + (r.flips ? L", longest " + MinSec (r.longest) : L"");
		return o;
		}
	}
