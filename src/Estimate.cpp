#include "Estimate.h"
#include "Csv.h"

#include <algorithm>
#include <cmath>

namespace
	{
	const double kPi = 3.14159265358979323846;

	double Rpm (const Estimate::Group &g, double dia)
		{
		if (!g.css)
			return g.speed;
		const double k = g.mm ? 1000.0 : 12.0;
		double rpm = dia > 1e-9 ? k * g.speed / (kPi * dia) : 1e12;
		if (g.cap > 0 && rpm > g.cap)
			rpm = g.cap;
		return rpm;
		}

	/// A number for a formula: exact enough, never in exponent form.
	std::wstring Num (double v)
		{
		return Csv::Tidy (v);
		}
	}

namespace Estimate
	{
	std::vector<Band> Bands (std::vector<std::pair<double, double>> pieces, size_t maxBands)
		{
		std::vector<Band> out;
		double total = 0;
		for (const auto &p : pieces)
			total += p.first;
		if (total <= 0 || maxBands == 0)
			return out;
		std::sort (pieces.begin (), pieces.end (),
				   [] (const std::pair<double, double> &a, const std::pair<double, double> &b)
				   { return std::fabs (a.second) < std::fabs (b.second); });
		const double share = total / static_cast<double> (maxBands);
		Band cur;
		for (const auto &p : pieces)
			{
			cur.len += p.first;
			cur.lenDia += p.first * 2.0 * std::fabs (p.second);
			if (cur.len >= share * (1.0 - 1e-9) && out.size () + 1 < maxBands)
				{
				out.push_back (cur);
				cur = Band ();
				}
			}
		if (cur.len > 0)
			out.push_back (cur);
		return out;
		}

	double Seconds (const Group &g)
		{
		const double f = std::fabs (g.feed);
		if (f <= 0)
			return 0;
		double s = 0;
		for (const Band &b : g.bands)
			{
			if (b.len <= 0)
				continue;
			if (g.feed < 0)
				{
				const double rpm = Rpm (g, b.lenDia / b.len);
				s += rpm > 0 ? 60.0 * b.len / (f * rpm) : 0;
				}
			else
				s += 60.0 * b.len / f;
			}
		return s;
		}

	std::wstring Term (const Group &g, const Cells &c)
		{
		double total = 0;
		for (const Band &b : g.bands)
			total += b.len;
		if (total <= 0 || g.feed == 0)
			return L"0";

		const std::wstring F = c.feed.empty () ? Num (std::fabs (g.feed)) : c.feed;
		const std::wstring S = c.speed.empty () ? Num (g.speed) : c.speed;
		const std::wstring cap = c.cap.empty () ? Num (g.cap) : c.cap;
		const std::wstring css = c.css.empty () ? (g.css ? L"TRUE" : L"FALSE")
							   : c.cssIsWord ? L"(" + c.css + L"=\"CSS\")" : L"(" + c.css + L"=1)";
		const std::wstring k = g.mm ? L"1000" : L"12";

		// per rev: 60/F * sum over bands of L / RPM(D)
		std::wstring perRev;
		for (const Band &b : g.bands)
			{
			if (b.len <= 0)
				continue;
			// On centre (a drill) the diameter is 0: CSS would ask for infinite RPM,
			// which the cap stops - never a division by zero.
			const std::wstring D = Num ((std::max) (1e-6, b.lenDia / b.len));
			const std::wstring surf = k + L"*" + S + L"/(PI()*" + D + L")";
			const std::wstring rpm = L"IF(" + css + L",IF(" + cap + L">0,MIN(" + cap + L"," + surf
									 + L")," + surf + L")," + S + L")";
			perRev += (perRev.empty () ? L"" : L"+") + Num (b.len) + L"/MAX(1E-9," + rpm + L")";
			}
		const std::wstring revSec = L"60/MAX(1E-12," + F + L")*(" + perRev + L")";
		const std::wstring minSec = L"60*" + Num (total) + L"/MAX(1E-12," + F + L")";

		if (!c.feedMode.empty ())
			return L"IF(" + c.feedMode + L"=\"per rev\"," + revSec + L"," + minSec + L")";
		return g.feed < 0 ? revSec : minSec;
		}
	}
