#include "Impact.h"

namespace
	{
	/// A column's place in a header row, or -1.
	int ColOf (const Csv::Row &head, const wchar_t *name)
		{
		for (size_t i = 0; i < head.size (); ++i)
			if (Csv::Trim (head[i]) == name)
				return static_cast<int> (i);
		return -1;
		}

	/// A cell as a number, NaN when it is empty or not one.
	double Num (const Csv::Row &row, int col)
		{
		double v = 0;
		if (col < 0 || static_cast<size_t> (col) >= row.size ()
			|| !Csv::ParseDouble (row[static_cast<size_t> (col)], v))
			return std::nan ("");
		return v;
		}

	/// The op_idn of a row, or 0.
	long OpOf (const Csv::Row &row, int col)
		{
		long long id = 0;
		if (col < 0 || static_cast<size_t> (col) >= row.size ()
			|| !Csv::ParseLong (row[static_cast<size_t> (col)], id) || id <= 0)
			return 0;
		return static_cast<long> (id);
		}

	/// Flips as a person counts them: whole where they are whole, else to the
	/// hundredth the sheet rounds its totals to.
	std::wstring Flips (double v)
		{
		return Csv::Tidy (std::round (v * 100.0) / 100.0);
		}

	const wchar_t *const kDot = L"  ·  ";
	}

namespace Impact
	{
	std::map<long, Op> Read (const std::vector<Csv::Row> &sheet,
							 const std::vector<Csv::Row> &dumped)
		{
		std::map<long, Op> ops;
		if (sheet.empty ())
			return ops;

		const Csv::Row &head = sheet[0];
		const int idCol = ColOf (head, L"op_idn"), rawCol = ColOf (head, L"cycle_time_raw"),
				  estCol = ColOf (head, L"est_seconds"), partCol = ColOf (head, L"flips_part"),
				  flipsCol = ColOf (head, L"flips");

		// Flips as dumped, by op, from the Dumped sheet's own flips_part.
		std::map<long, double> dumpedFlips;
		if (!dumped.empty ())
			{
			const int dId = ColOf (dumped[0], L"op_idn"), dPart = ColOf (dumped[0], L"flips_part");
			for (size_t k = 1; k < dumped.size (); ++k)
				{
				const long op = OpOf (dumped[k], dId);
				const double v = Num (dumped[k], dPart);
				if (op != 0 && !std::isnan (v))
					dumpedFlips[op] = v;
				}
			}
		const bool haveDumped = !dumpedFlips.empty ();

		for (size_t k = 1; k < sheet.size (); ++k)
			{
			const long op = OpOf (sheet[k], idCol);
			if (op == 0)
				continue;
			Op o;
			o.op = op;
			o.was = Num (sheet[k], rawCol);
			o.now = Num (sheet[k], estCol);
			o.flipsNow = Num (sheet[k], partCol);
			if (haveDumped)
				{
				const auto at = dumpedFlips.find (op);
				if (at != dumpedFlips.end ())
					o.flipsWas = at->second;
				}
			else
				o.flipsWas = Num (sheet[k], flipsCol);
			ops[op] = o;
			}
		return ops;
		}

	Total Sum (const std::map<long, Op> &ops, const std::set<long> &applied)
		{
		Total t;
		for (const auto &kv : ops)
			{
			const Op &o = kv.second;
			const bool on = applied.count (o.op) != 0;
			const double now = on ? o.now : o.was;
			const double flipsNow = on ? o.flipsNow : o.flipsWas;
			if (!std::isnan (o.was) && !std::isnan (now))
				{
				t.was += o.was;
				t.now += now;
				++t.timed;
				}
			if (!std::isnan (o.flipsWas) && !std::isnan (flipsNow))
				{
				t.flipsWas += o.flipsWas;
				t.flipsNow += flipsNow;
				++t.flipped;
				}
			}
		return t;
		}

	std::wstring Hms (double seconds)
		{
		if (std::isnan (seconds) || seconds < 0)
			seconds = 0;
		const long long s = static_cast<long long> (std::llround (seconds));
		wchar_t buf[48];
		swprintf (buf, 48, L"%lld:%02lld:%02lld", s / 3600, (s / 60) % 60, s % 60);
		return buf;
		}

	std::wstring Change (double seconds)
		{
		// Rounded first, so a change of 0.4 s reads +0:00:00 and never -0:00:00.
		const double r = std::round (seconds);
		return (r < 0 ? L"-" : L"+") + Hms (std::fabs (r));
		}

	int Tone (double was, double now, bool flips)
		{
		if (std::isnan (was) || std::isnan (now))
			return 0;
		const double d = flips ? std::round ((now - was) * 100.0) : std::round (now) - std::round (was);
		return d < 0 ? -1 : d > 0 ? 1 : 0;
		}

	std::wstring OpText (const Op &o)
		{
		std::wstring s;
		if (Tone (o.was, o.now) != 0)
			s = Hms (o.was) + L" -> " + Hms (o.now) + L" (" + Change (std::round (o.now) - std::round (o.was)) + L")";
		if (Tone (o.flipsWas, o.flipsNow, true) != 0)
			s += (s.empty () ? L"" : kDot) + std::wstring (L"flips ") + Flips (o.flipsWas) + L" -> " + Flips (o.flipsNow);
		return s;
		}

	std::wstring TotalText (const Total &t)
		{
		std::wstring s;
		if (t.timed > 0)
			s = Tone (t.was, t.now) == 0
					? L"Cycle time " + Hms (t.was) + L", unchanged"
					: L"Cycle time " + Hms (t.was) + L" -> " + Hms (t.now) + L" ("
						  + Change (std::round (t.now) - std::round (t.was)) + L")";
		if (t.flipped > 0)
			s += (s.empty () ? std::wstring (L"Flips ") : kDot + std::wstring (L"flips "))
				 + (Tone (t.flipsWas, t.flipsNow, true) == 0
						? Flips (t.flipsWas) + L", unchanged"
						: Flips (t.flipsWas) + L" -> " + Flips (t.flipsNow));
		return s;
		}
	}
