// Mill MRR (ae x ap x feed per minute, the number and its live formula) and
// transforms in words. Rows are built the way the dump fills them: each cell's
// text as the sheet shows it, and where it sits.
#include "../src/MillMrr.h"
#include "../src/Xform.h"

#include <cstdio>
#include <cmath>

static int failed = 0;
static void Check (bool ok, const char *what)
	{
	std::printf ("  %s  %s\n", ok ? "ok  " : "FAIL", what);
	if (!ok)
		++failed;
	}

static bool Has (const std::wstring &s, const std::wstring &part)
	{
	return s.find (part) != std::wstring::npos;
	}

/// A row: column name -> text, each at its own made-up cell (C3, D3 ...).
static MillMrr::Cells Row (std::initializer_list<std::pair<const wchar_t *, const wchar_t *>> cells)
	{
	MillMrr::Cells out;
	wchar_t col = L'C';
	for (const auto &c : cells)
		out[c.first] = { c.second, std::wstring (1, col++) + L"3" };
	return out;
	}

int main ()
	{
	// Dynamic mill, inch: 0.1 stepover, the whole 0.5 depth, 100 in/min.
	MillMrr::Cells dyn = Row ({ { L"speed", L"8000" }, { L"speed_mode", L"RPM" }, { L"feed", L"100" },
								{ L"feed_mode", L"per min" }, { L"stepover", L"0.1" }, { L"dcuts_on", L"0" },
								{ L"dcut_rough", L"0.2" }, { L"depth", L"-0.5" }, { L"depth_inc", L"0" },
								{ L"top_stock", L"0" }, { L"top_stock_inc", L"0" } });
	MillMrr::Result r = MillMrr::Compute (L"DYNAMIC MILL", false, 0.5, dyn);
	Check (r.ok && r.mrr == 5.0, "dynamic: 0.1 x 0.5 x 100 = 5 in3/min");
	Check (Has (r.formula, L"MIN(") && Has (r.formula, L",0.5)") && r.formula.rfind (L"IFERROR(ROUND(", 0) == 0,
		   "dynamic: stepover capped at the tool diameter, in IFERROR(ROUND(...))");
	Check (Has (r.basis, L"ae = stepover 0.1") && Has (r.basis, L"whole depth 0.5") && Has (r.basis, L"in3/min"),
		   "dynamic: basis names ae, ap and units");

	// Depth cuts on: the rough step, never more than the whole depth.
	dyn[L"dcuts_on"].text = L"1";
	r = MillMrr::Compute (L"DYNAMIC MILL", false, 0.5, dyn);
	Check (r.ok && r.mrr == 2.0, "dynamic, depth cuts 0.2: 0.1 x 0.2 x 100 = 2");
	dyn[L"dcut_rough"].text = L"0.8";
	r = MillMrr::Compute (L"DYNAMIC MILL", false, 0.5, dyn);
	Check (r.ok && r.mrr == 5.0, "a rough step past the whole depth is the whole depth");

	// Per rev feed: x RPM.
	MillMrr::Cells slot = Row ({ { L"speed", L"6000" }, { L"speed_mode", L"RPM" }, { L"feed", L"0.002" },
								 { L"feed_mode", L"per rev" }, { L"fr_override_on", L"0" }, { L"fr_override", L"30" },
								 { L"dcuts_on", L"0" }, { L"dcut_rough", L"0.1" }, { L"mcuts_on", L"0" },
								 { L"mcut_rough_n", L"2" }, { L"mcut_rough_amt", L"0.05" }, { L"depth", L"-0.125" },
								 { L"depth_inc", L"1" }, { L"top_stock", L"0" }, { L"top_stock_inc", L"1" } });
	r = MillMrr::Compute (L"CONTOUR", false, 0.25, slot);
	Check (r.ok && r.mrr == 0.375, "contour full slot: 0.25 x 0.125 x (0.002 x 6000) = 0.375");
	Check (Has (r.basis, L"full slot") && Has (r.basis, L"per rev x RPM"), "contour: says full slot, per rev x RPM");
	Check (Has (r.formula, L"\"per rev\""), "contour: the formula follows feed_mode");

	slot[L"mcuts_on"].text = L"1";
	r = MillMrr::Compute (L"CONTOUR", false, 0.25, slot);
	Check (r.ok && r.mrr == 0.075, "contour multi passes: 0.05 x 0.125 x 12 = 0.075");
	Check (Has (r.basis, L"multi-pass spacing 0.05"), "contour multi passes: said so");

	slot[L"fr_override_on"].text = L"1";
	r = MillMrr::Compute (L"CONTOUR", false, 0.25, slot);
	Check (r.ok && std::fabs (r.mrr - 0.188) < 1e-12, "feed rate override 30: 0.05 x 0.125 x 30 = 0.1875 -> 0.188");
	Check (Has (r.basis, L"feed rate override"), "override: said so");

	// CSS on a live tool: surface speed at the tool's diameter, capped.
	MillMrr::Cells css = Row ({ { L"speed", L"500" }, { L"speed_mode", L"CSS" }, { L"max_ss", L"3000" },
								{ L"feed", L"0.004" }, { L"feed_mode", L"per rev" }, { L"stepover", L"0.05" },
								{ L"depth", L"-1" }, { L"depth_inc", L"0" }, { L"top_stock", L"0" },
								{ L"top_stock_inc", L"0" } });
	r = MillMrr::Compute (L"DYNAMIC MILL", false, 0.5, css);
	// 500 SFM at 0.5" = 3819.7 RPM, capped at 3000: 0.05 x 1 x 0.004 x 3000 = 0.6
	Check (r.ok && r.mrr == 0.6, "CSS capped by max_ss: 0.05 x 1 x 12 = 0.6");

	// Metric: mm x mm x mm/min / 1000 = cm3/min.
	MillMrr::Cells mm = Row ({ { L"speed", L"5000" }, { L"speed_mode", L"RPM" }, { L"feed", L"1000" },
							   { L"feed_mode", L"per min" }, { L"stepover", L"2" }, { L"depth", L"-5" },
							   { L"depth_inc", L"0" }, { L"top_stock", L"0" }, { L"top_stock_inc", L"0" } });
	r = MillMrr::Compute (L"DYNAMIC MILL", true, 10, mm);
	Check (r.ok && r.mrr == 10.0, "metric: 2 x 5 x 1000 / 1000 = 10 cm3/min");
	Check (Has (r.formula, L"/1000") && Has (r.basis, L"cm3/min"), "metric: /1000 and cm3/min");

	// Nothing to go on: no depth (absolute top, incremental depth), no depth cuts.
	mm[L"depth_inc"].text = L"1";
	Check (!MillMrr::Compute (L"DYNAMIC MILL", true, 10, mm).ok, "depth and top of stock measured differently: blank");
	Check (!MillMrr::Compute (L"ROUGH", false, 0.5, dyn).ok, "a lathe kind is not milling");
	Check (!MillMrr::Compute (L"DYNAMIC MILL", false, 0, dyn).ok, "no tool diameter: blank");

	// ---- Transforms.
	Xform::Params rot;
	rot.type = 2;
	rot.rotSteps = 4;
	rot.rotAngle = 90;
	Check (Xform::KindName (2) == L"rotate" && Xform::KindName (9) == L"9", "kind names");
	Check (Xform::Instances (rot) == 4, "rotate: its steps");
	Check (Xform::Describe (rot) == L"rotate 4 x 90 deg from 0 deg about (0, 0, 0)", "rotate in words");

	Xform::Params rect;
	rect.type = 3;
	rect.trnStyle = 17;
	rect.trnSteps[0] = 3;
	rect.trnSteps[1] = 2;
	rect.trnDist[0] = 1.5;
	rect.trnDist[1] = 2;
	Check (Xform::Instances (rect) == 6, "rectangular 3 x 2: 6");
	Check (Xform::Describe (rect) == L"translate rectangular 3 x 2, X 1.5 Y 2 apart", "rectangular in words");
	rect.trnStyle = 18;
	Check (Xform::Instances (rect) == 3, "polar: one direction only");

	Xform::Params mir;
	mir.type = 1;
	mir.mirrorTo[0] = 1;
	Check (Xform::Instances (mir) == 1, "mirror: one copy");
	Check (Xform::Describe (mir) == L"mirror about (0, 0, 0) - (1, 0, 0)", "mirror in words");
	Check (Xform::IdList ({ 12, 13, 20 }) == L"12, 13, 20" && Xform::IdList ({}).empty (), "op id list");

	if (failed)
		std::printf ("mill_test: %d FAILED\n", failed);
	else
		std::printf ("mill_test: all checks passed\n");
	return failed ? 1 : 0;
	}
