#include "Xform.h"
#include "Csv.h"

#include <cmath>

namespace Xform
	{
	namespace
		{
		/// A number as a person reads it - 4 places at most.
		std::wstring N (double v)
			{
			return Csv::Tidy (std::round (v * 10000.0) / 10000.0);
			}

		std::wstring Point (const double p[3])
			{
			return L"(" + N (p[0]) + L", " + N (p[1]) + L", " + N (p[2]) + L")";
			}
		}

	std::wstring KindName (int type)
		{
		switch (type)
			{
			case 1: return L"mirror";
			case 2: return L"rotate";
			case 3: return L"translate";
			}
		return std::to_wstring (type);
		}

	long Instances (const Params &p)
		{
		switch (p.type)
			{
			case 1:
				return 1;
			case 2:
				return p.rotSteps > 0 ? p.rotSteps : 0;
			case 3:
				{
				const long x = p.trnSteps[0] > 0 ? p.trnSteps[0] : 0;
				// Only the rectangular array has a second direction.
				if (p.trnStyle == 17)
					return x * (p.trnSteps[1] > 0 ? p.trnSteps[1] : 1);
				return x;
				}
			}
		return 0;
		}

	std::wstring Describe (const Params &p)
		{
		switch (p.type)
			{
			case 1:
				return L"mirror about " + Point (p.mirrorFrom) + L" - " + Point (p.mirrorTo);
			case 2:
				return L"rotate " + std::to_wstring (p.rotSteps) + L" x "
					   + (p.rotTotal ? L"(total " + N (p.rotAngle) + L" deg)" : N (p.rotAngle) + L" deg")
					   + L" from " + N (p.rotStart) + L" deg about " + Point (p.rotAbout);
			case 3:
				{
				const std::wstring how = p.trnTotal ? L" total" : L" apart";
				switch (p.trnStyle)
					{
					case 17:
						return L"translate rectangular " + std::to_wstring (p.trnSteps[0]) + L" x "
							   + std::to_wstring (p.trnSteps[1]) + L", X " + N (p.trnDist[0]) + L" Y "
							   + N (p.trnDist[1]) + how + (p.zigzag ? L", zigzag" : L"");
					case 18:
						return L"translate polar " + std::to_wstring (p.trnSteps[0]) + L" x " + N (p.trnPolar[0])
							   + how + L" at " + N (p.trnPolar[1]) + L" deg";
					case 19:
						return L"translate between points, " + std::to_wstring (p.trnSteps[0]) + L" step(s)";
					case 20:
						return L"translate between views, " + std::to_wstring (p.trnSteps[0]) + L" step(s)";
					}
				return L"translate (style " + std::to_wstring (p.trnStyle) + L"), "
					   + std::to_wstring (p.trnSteps[0]) + L" step(s)";
				}
			}
		return L"transform type " + std::to_wstring (p.type);
		}

	std::wstring IdList (const std::vector<long> &ids)
		{
		std::wstring s;
		for (long id : ids)
			s += (s.empty () ? L"" : L", ") + std::to_wstring (id);
		return s;
		}
	}
