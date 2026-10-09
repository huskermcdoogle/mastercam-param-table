//
// Xform.h - a TRANSFORM operation in words: what kind, how many copies, and
// of which operations. No Mastercam SDK - the dump copies the few numbers out
// of the operation's prm_xform and hands them here.
//
// A transform is shown, never edited: its sources and instances are worked
// out by Mastercam from geometry and other operations, not from one number a
// sheet could change safely.
//
#pragma once

#include <string>
#include <vector>

namespace Xform
	{
	/// The numbers that matter from prm_xform. Only the member of the union
	/// for `type` is meaningful - the others hold whatever the union does.
	struct Params
		{
		int type = 0;				//!< 1 mirror, 2 rotate, 3 translate
		// Mirror
		double mirrorFrom[3] = { 0, 0, 0 }, mirrorTo[3] = { 0, 0, 0 };
		// Rotate
		int rotSteps = 0;			//!< n_steps
		double rotAngle = 0;		//!< degrees: between steps, or the total when rotTotal
		double rotStart = 0;		//!< degrees
		bool rotTotal = false;		//!< distMode 1
		double rotAbout[3] = { 0, 0, 0 };
		// Translate
		int trnStyle = 0;			//!< pt_type: 17 rectangular, 18 polar, 19 between points, 20 between views
		int trnSteps[2] = { 0, 0 };
		double trnDist[2] = { 0, 0 };
		bool trnTotal = false;		//!< distMode 1
		double trnPolar[2] = { 0, 0 };	//!< distance, angle in degrees
		bool zigzag = false;
		};

	/// "mirror", "rotate", "translate", or the number when it is none of them.
	std::wstring KindName (int type);

	/// How many copies the transform makes (not counting the original); 0 when
	/// the parameters do not say.
	long Instances (const Params &p);

	/// One line: the kind and its numbers ("rotate 4 x 90 deg from 0 deg about
	/// (0, 0, 0)").
	std::wstring Describe (const Params &p);

	/// Op ids as "12, 13, 14".
	std::wstring IdList (const std::vector<long> &ids);
	}
