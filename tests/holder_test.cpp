// The ISO / ANSI code readers (ToolPictures.h, no SDK): a holder code's style
// letter -> its entering angle, and an insert code -> its size (IC). The holder
// names are ones Mastercam's own, Sandvik's and Kennametal's lathe libraries use.
#include "../src/ToolPictures.h"

#include <cmath>
#include <cstdio>

static int failed = 0;
static void Check (bool ok, const char *what)
	{
	std::printf ("  %s  %s\n", ok ? "ok  " : "FAIL", what);
	if (!ok)
		++failed;
	}

static bool Near (double a, double b)
	{
	return std::fabs (a - b) < 1e-6;
	}

int main ()
	{
	using namespace ToolPictures;

	// ---- The style table: ISO 5608 table 3 (and Q from ISO 6261).
	Check (StyleAngle (L'A') == 90 && StyleAngle (L'B') == 75 && StyleAngle (L'C') == 90 && StyleAngle (L'D') == 45
			   && StyleAngle (L'E') == 60 && StyleAngle (L'F') == 90 && StyleAngle (L'G') == 90,
		   "styles A-G: 90 75 90 45 60 90 90");
	Check (StyleAngle (L'H') == 107.5 && StyleAngle (L'J') == 93 && StyleAngle (L'K') == 75 && StyleAngle (L'L') == 95
			   && StyleAngle (L'M') == 50 && StyleAngle (L'N') == 63,
		   "styles H-N: 107.5 93 75 95 50 63");
	Check (StyleAngle (L'P') == 117.5 && StyleAngle (L'Q') == 107.5 && StyleAngle (L'R') == 75 && StyleAngle (L'S') == 45
			   && StyleAngle (L'T') == 60 && StyleAngle (L'U') == 93 && StyleAngle (L'V') == 72.5 && StyleAngle (L'W') == 60
			   && StyleAngle (L'Y') == 85,
		   "styles P-Y: 117.5 107.5 75 45 60 93 72.5 60 85");
	Check (StyleAngle (L'I') == 0 && StyleAngle (L'O') == 0 && StyleAngle (L'X') == 0 && StyleAngle (L'b') == 0 && StyleAngle (0) == 0,
		   "no style: I, O, X, Mastercam's button 'b', none");

	// ---- Holder codes in names.
	Check (HolderCode (L"PCLNR 2525M 12") == L"PCLNR" && HolderAngle (L"PCLNR 2525M 12") == 95, "Sandvik PCLNR 2525M 12: L, 95");
	Check (HolderCode (L"PCLNL2020K12") == L"PCLNL", "PCLNL2020K12 run together");
	Check (HolderAngle (L"MCLNR-164D") == 95, "ANSI MCLNR-164D: 95");
	Check (HolderAngle (L"MVJNR-164D") == 93 && HolderAngle (L"MDJNL-164C") == 93 && HolderAngle (L"MTJNL 2525M 16") == 93,
		   "J style (MVJNR, MDJNL, MTJNL): 93");
	Check (HolderAngle (L"PSSNR 3225P 15") == 45 && HolderAngle (L"CSDNN- R-164") == 0, "PSSNR: 45; CSDNN- R-164 (no figure after the dash): none");
	Check (HolderAngle (L"DCMNN-164C") == 50 && HolderAngle (L"DTANLS-164") == 90 && HolderAngle (L"CTENR-102") == 60,
		   "Kennametal DCMNN 50, DTANLS (sixth letter) 90, CTENR 60");
	Check (HolderAngle (L"DTJNLS-164D") == 93, "Kennametal DTJNLS-164D: 93");
	Check (HolderAngle (L"S25T-PCLNR12") == 95 && HolderAngle (L"S25T-SDUCR 11") == 93 && HolderAngle (L"A20-MDUNR4") == 93,
		   "boring bars after their dash: S25T-PCLNR12 95, S25T-SDUCR 11 93, A20-MDUNR4 93");
	Check (HolderAngle (L"S16R-STFCL 11") == 90 && HolderAngle (L"A20-DDQNL4") == 107.5 && HolderAngle (L"A20-DDPNR4") == 117.5,
		   "boring bars F 90, Q 107.5, P 117.5");
	Check (HolderAngle (L"PSBNR 2525M 12") == 75 && HolderAngle (L"MVVNN 2525M 16") == 72.5 && HolderAngle (L"PWLNR 2020K 08") == 95,
		   "PSBNR 75, MVVNN 72.5, PWLNR 95");
	Check (HolderAngle (L"OD ROUGH MCLNR-164D CNMG 432") == 95, "a code inside the tool's name");
	Check (HolderAngle (L"SRGCR 2525M 10") == 0 && HolderCode (L"SRGCR 2525M 10") == L"SRGCR", "a round insert's holder (SRGCR): no angle");
	Check (HolderAngle (L"STEEL 4140") == 0, "not a holder: STEEL 4140");
	Check (HolderAngle (L"CNMG 432") == 0 && HolderAngle (L"CNMG120408") == 0, "not a holder: an insert code");
	Check (HolderAngle (L"OD ROUGH RIGHT - 80 DEG.") == 0 && HolderAngle (L"") == 0, "not a holder: words, nothing");
	Check (HolderAngle (L"XPCLNR 2525") == 0, "not a holder: letters glued in front");
	Check (HolderAngle (L"MCLNR") == 0, "not a holder: no size figures after the letters");
	Check (HolderAngle (L"CP-30AR-2020-11") == 0 && HolderAngle (L"C4-CP-30AL-27050-11C-60") == 0, "PrimeTurning holders: no angle");

	// ---- Insert codes -> IC.
	Check (Near (InsertCodeIC (L"CNMG 432", false), 0.5) && Near (InsertCodeIC (L"CNMG-432", false), 0.5)
			   && Near (InsertCodeIC (L"CNMG432", false), 0.5),
		   "CNMG 432 / -432 / 432: 0.5 in");
	Check (Near (InsertCodeIC (L"CNMG 432", true), 12.7), "CNMG 432 in a metric part: 12.7 mm");
	Check (Near (InsertCodeIC (L"TPMT-21.51LF", false), 0.25) && Near (InsertCodeIC (L"VNMG-332", false), 0.375),
		   "TPMT-21.51LF 0.25, VNMG-332 0.375");
	Check (Near (InsertCodeIC (L"SNG-422", false), 0.5), "old three-letter SNG-422: 0.5");
	Check (Near (InsertCodeIC (L"CNMG 12 04 08", true), 12.7) && Near (InsertCodeIC (L"CNMG120408", true), 12.7)
			   && Near (InsertCodeIC (L"CNMG 120408-PM", true), 12.7),
		   "CNMG 12 04 08 / 120408 / 120408-PM: 12.7 mm");
	Check (Near (InsertCodeIC (L"CNMG 120408", false), 0.5), "CNMG 120408 in an inch part: 0.5 in");
	Check (Near (InsertCodeIC (L"CCMT 060204", true), 6.35) && Near (InsertCodeIC (L"CNMG 090304", true), 9.525)
			   && Near (InsertCodeIC (L"CNMG 160608", true), 15.875) && Near (InsertCodeIC (L"CNMG 190608", true), 19.05),
		   "C 06 6.35, 09 9.525, 16 15.875, 19 19.05");
	Check (Near (InsertCodeIC (L"DCMT 11T304", true), 9.525) && Near (InsertCodeIC (L"DNMG 150608", true), 12.7)
			   && Near (InsertCodeIC (L"DCMT 070204", true), 6.35),
		   "D 11T3 9.525, 15 12.7, 07 6.35");
	Check (Near (InsertCodeIC (L"VBMT 16 04 04", true), 9.525) && Near (InsertCodeIC (L"VCMT 110302", true), 6.35),
		   "V 16 9.525, 11 6.35");
	Check (Near (InsertCodeIC (L"TNMG 160408", true), 9.525) && Near (InsertCodeIC (L"TCMT 110204", true), 6.35)
			   && Near (InsertCodeIC (L"TNMG 220408", true), 12.7),
		   "T 16 9.525, 11 6.35, 22 12.7");
	Check (Near (InsertCodeIC (L"SNMG 120408", true), 12.7) && Near (InsertCodeIC (L"SNMG 090308", true), 9.525),
		   "S 12 12.7, 09 9.525");
	Check (Near (InsertCodeIC (L"WNMG 080408", true), 12.7) && Near (InsertCodeIC (L"WNMG 06 04 08", true), 9.525),
		   "W 08 12.7, 06 9.525");
	Check (Near (InsertCodeIC (L"RCMT 10T3MO", true), 10) && Near (InsertCodeIC (L"RCMT 10 T3 M0", true), 10)
			   && Near (InsertCodeIC (L"RCMT 1606MO", true), 16),
		   "round RCMT 10T3MO / 10 T3 M0: 10 mm; 1606: 16");
	Check (Near (InsertCodeIC (L"RCMT 1606MO", false), 0.6299), "a 16 mm round in an inch part: 0.6299");
	Check (Near (InsertCodeIC (L"RCMT-32", false), 0) && Near (InsertCodeIC (L"RCMT-32.5", false), 0.375),
		   "inch round RCMT-32.5: 0.375 (a lone 32 is too short to trust)");
	Check (Near (InsertCodeIC (L"OD ROUGH CNMG 432 R", false), 0.5), "a code inside a name");
	Check (InsertCodeIC (L"MCLNR-164D", false) == 0 && InsertCodeIC (L"SCLCR 12", true) == 0 && InsertCodeIC (L"DCLNR-164C", false) == 0,
		   "a holder code is not an insert code");
	Check (InsertCodeIC (L"CNC 4", false) == 0 && InsertCodeIC (L"TEST 1", false) == 0 && InsertCodeIC (L"", false) == 0,
		   "not codes: CNC 4, TEST 1, nothing");
	Check (InsertCodeIC (L"KNUX 160405", true) == 0, "a K (parallelogram): no size from the code");

	if (failed)
		std::printf ("holder_test: %d FAILED\n", failed);
	else
		std::printf ("holder_test: all checks passed\n");
	return failed ? 1 : 0;
	}
