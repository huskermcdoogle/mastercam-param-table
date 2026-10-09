#include "MillMrr.h"
#include "Csv.h"

#include <algorithm>
#include <cmath>

namespace MillMrr
	{
	Result Compute (const std::wstring &type, bool mm, double toolDia, const Cells &cells)
		{
		Result r;
		const bool contour = type == L"CONTOUR", dynamic = type == L"DYNAMIC MILL";
		if (!contour && !dynamic)
			return r;

		auto has = [&cells] (const wchar_t *name) { return cells.count (name) > 0; };
		auto text = [&cells] (const wchar_t *name)
			{
			const auto it = cells.find (name);
			return it == cells.end () ? std::wstring () : it->second.text;
			};
		auto ref = [&cells] (const wchar_t *name)
			{
			const auto it = cells.find (name);
			return it == cells.end () ? std::wstring () : it->second.ref;
			};
		auto num = [&] (const wchar_t *name)
			{
			double v = 0;
			Csv::ParseDouble (text (name), v);
			return v;
			};
		// A switch cell holds 1 or 0 - and the formulas test it the same way (=1).
		auto on = [&] (const wchar_t *name) { return has (name) && num (name) == 1.0; };

		if (!has (L"speed") || !has (L"feed") || !has (L"feed_mode"))
			return r;

		// The diameter goes into the formula as a number - so the cached value is
		// worked from that same number, not from the unrounded one.
		const std::wstring D = Csv::Tidy (toolDia);
		double dia = 0;
		if (!Csv::ParseDouble (D, dia) || !(dia > 0))
			return r;
		const double pi = 3.14159265358979323846;

		// ---- The spindle. A milling tool runs at RPM; CSS (rare on a live tool)
		// is taken at the tool's own diameter, capped by max_ss.
		const std::wstring K = mm ? L"1000" : L"12";
		const double kSpeed = mm ? 1000.0 : 12.0;
		const bool css = has (L"speed_mode") && text (L"speed_mode") == L"CSS";
		const double speed = num (L"speed"), cap = has (L"max_ss") ? num (L"max_ss") : 0;
		double rpm = css ? speed * kSpeed / (pi * dia) : speed;
		if (css && cap > 0)
			rpm = (std::min) (rpm, cap);
		std::wstring rpmF = ref (L"speed");
		if (has (L"speed_mode"))
			{
			const std::wstring capRef = has (L"max_ss") ? ref (L"max_ss") : L"0";
			const std::wstring fromSurface = ref (L"speed") + L"*" + K + L"/(PI()*" + D + L")";
			rpmF = L"IF(" + ref (L"speed_mode") + L"=\"CSS\",IF(N(" + capRef + L")>0,MIN(" + fromSurface + L","
				   + capRef + L")," + fromSurface + L")," + ref (L"speed") + L")";
			}

		// ---- Feed per minute. The contour's feed rate override, when it is on,
		// is the feed the cutting moves are posted at.
		const bool perRev = text (L"feed_mode") == L"per rev";
		double fpm = perRev ? num (L"feed") * rpm : num (L"feed");
		std::wstring fpmF = L"IF(" + ref (L"feed_mode") + L"=\"per rev\"," + ref (L"feed") + L"*" + rpmF + L","
							+ ref (L"feed") + L")";
		bool overridden = false;
		if (has (L"fr_override_on") && has (L"fr_override"))
			{
			overridden = on (L"fr_override_on") && num (L"fr_override") > 0;
			if (overridden)
				fpm = num (L"fr_override");
			fpmF = L"IF(AND(" + ref (L"fr_override_on") + L"=1,N(" + ref (L"fr_override") + L")>0),"
				   + ref (L"fr_override") + L"," + fpmF + L")";
			}
		if (!(fpm > 0))
			return r;

		// ---- Axial depth, ap. The whole depth is top of stock to depth - only
		// when both are given the same way (both absolute or both incremental);
		// one of each is measured from different places and is not a depth.
		double total = 0;
		std::wstring totalF;
		if (has (L"depth") && has (L"top_stock") && has (L"depth_inc") && has (L"top_stock_inc")
			&& text (L"depth_inc") == text (L"top_stock_inc"))
			{
			total = std::fabs (num (L"top_stock") - num (L"depth"));
			if (total > 0)
				totalF = L"ABS(" + ref (L"top_stock") + L"-" + ref (L"depth") + L")";
			}
		double ap = 0;
		std::wstring apF, apWords;
		const bool dcutCols = has (L"dcuts_on") && has (L"dcut_rough");
		const bool dcuts = dcutCols && on (L"dcuts_on") && num (L"dcut_rough") > 0;
		if (dcuts)
			{
			// The rough step is the most each depth cut takes; never more than the whole depth.
			ap = totalF.empty () ? num (L"dcut_rough") : (std::min) (num (L"dcut_rough"), total);
			apWords = L"ap = depth cut " + Csv::Tidy (ap);
			}
		else if (!totalF.empty ())
			{
			ap = total;
			apWords = L"ap = whole depth " + Csv::Tidy (ap) + L" (top_stock to depth)";
			}
		else
			return r;
		const std::wstring roughF = totalF.empty () ? ref (L"dcut_rough")
													: L"MIN(" + ref (L"dcut_rough") + L"," + totalF + L")";
		if (dcutCols)
			apF = L"IF(" + ref (L"dcuts_on") + L"=1," + roughF + L"," + (totalF.empty () ? L"\"\"" : totalF) + L")";
		else
			apF = totalF;

		// ---- Radial engagement, ae. Never wider than the tool.
		double ae = 0;
		std::wstring aeF, aeWords;
		if (dynamic)
			{
			if (!has (L"stepover") || !(num (L"stepover") > 0))
				return r;
			ae = (std::min) (num (L"stepover"), dia);
			aeF = L"MIN(" + ref (L"stepover") + L"," + D + L")";
			aeWords = L"ae = stepover " + Csv::Tidy (ae);
			}
		else
			{
			const bool mcCols = has (L"mcuts_on") && has (L"mcut_rough_n") && has (L"mcut_rough_amt");
			if (mcCols && on (L"mcuts_on") && num (L"mcut_rough_n") > 0 && num (L"mcut_rough_amt") > 0)
				{
				ae = (std::min) (num (L"mcut_rough_amt"), dia);
				aeWords = L"ae = multi-pass spacing " + Csv::Tidy (ae);
				}
			else
				{
				// Nothing in a contour's parameters says how much of the tool's width
				// is in the cut: a full slot is the most it can be - an upper bound.
				ae = dia;
				aeWords = L"ae = tool dia " + D + L" (full slot - an upper bound)";
				}
			aeF = mcCols ? L"IF(AND(" + ref (L"mcuts_on") + L"=1,N(" + ref (L"mcut_rough_n") + L")>0,N("
							   + ref (L"mcut_rough_amt") + L")>0),MIN(" + ref (L"mcut_rough_amt") + L"," + D + L"),"
							   + D + L")"
						 : D;
			}

		double mrr = ae * ap * fpm / (mm ? 1000.0 : 1.0);
		r.mrr = std::round (mrr * 1000.0) / 1000.0;
		r.formula = L"IFERROR(ROUND(" + aeF + L"*" + apF + L"*" + fpmF + (mm ? L"/1000" : L"") + L",3),\"\")";
		r.basis = aeWords + L" x " + apWords + L" x feed " + Csv::Tidy (std::round (fpm * 1000.0) / 1000.0)
				  + L" per min" + (overridden ? L" (feed rate override)" : perRev ? L" (per rev x RPM)" : L"")
				  + (mm ? L" - cm3/min" : L" - in3/min");
		r.ok = true;
		return r;
		}
	}
