#include "stdafx.h"
#include "MastercamSdk.h"
#include "LatheFields.h"
#include "Coolant.h"
#include "Inspect.h"
#include "Csv.h"

#include <algorithm>
#include <cmath>

namespace Lathe
	{
	namespace
		{
		using Acc = void *(*) (void *);

		/// The address of one member of one parameter struct.
		#define ACC(S, MEMBER)                                              \
			[] (void *p) -> void *                                          \
				{ return static_cast<void *> (&static_cast<S *> (p)->MEMBER); }

		/// The address of one member of the OPERATION itself (feeds and speeds,
		/// home position, planes, reference points) rather than of its
		/// toolpath-specific struct.
		#define ACCO(MEMBER)                                                \
			[] (void *p) -> void *                                          \
				{ return static_cast<void *> (&static_cast<operation *> (p)->MEMBER); }

		const double kNone = std::nan ("");

		void AddImpl (Table &t, Base base, const wchar_t *name, Kind kind, Acc at,
					  bool readOnly, double lo, double hi, size_t textLen)
			{
			Plan::Col c;
			c.name = name;
			c.readOnly = readOnly;
			c.lo = lo;
			c.hi = hi;
			switch (kind)
				{
				case Kind::Bool:   c.type = Plan::Type::Bool;   break;
				case Kind::Double: c.type = Plan::Type::Double; break;
				case Kind::Coolant:
					c.type = Plan::Type::Text;
					c.hi = 400.0;
					c.check = &Coolant::Check;
					break;
				case Kind::CoolantV9:
					c.type = Plan::Type::Text;
					c.choices = Coolant::ChoicesV9 ();
					break;
				case Kind::MinSec:
					c.type = Plan::Type::Text;
					c.hi = 12.0;
					c.check = &Inspect::CheckMinSec;
					break;
				case Kind::DoubleSize: c.type = Plan::Type::Double; break;
				case Kind::LongSize:   c.type = Plan::Type::Long;   break;
				case Kind::DoubleSign:
				case Kind::LongSign:
				case Kind::BoolWord:
				case Kind::ShortWord:  c.type = Plan::Type::Text;   break;
				case Kind::Text:
					c.type = Plan::Type::Text;
					// The buffer holds textLen characters INCLUDING the NUL.
					c.hi = static_cast<double> (textLen) - 1.0;
					break;
				default:           c.type = Plan::Type::Long;   break;
				}
			t.schema.cols.push_back (c);

			Binding b;
			b.kind = kind;
			b.base = base;
			b.at = at;
			b.textLen = textLen;
			t.bindings.push_back (b);
			}

		/// A column in the toolpath-specific struct.
		void Add (Table &t, const wchar_t *name, Kind kind, Acc at,
				  bool readOnly = false, double lo = kNone, double hi = kNone,
				  size_t textLen = 0)
			{
			AddImpl (t, Base::Prm, name, kind, at, readOnly, lo, hi, textLen);
			}

		/// A column on the operation itself.
		void AddOp (Table &t, const wchar_t *name, Kind kind, Acc at,
					bool readOnly = false, double lo = kNone, double hi = kNone)
			{
			AddImpl (t, Base::Op, name, kind, at, readOnly, lo, hi, 0);
			}

		/// The words of the column just added: for a sign, [negative, positive];
		/// for a bool, [false, true]. They are also its only allowed values.
		void Words (Table &t, const wchar_t *first, const wchar_t *second)
			{
			t.bindings.back ().words[0] = first;
			t.bindings.back ().words[1] = second;
			t.schema.cols.back ().choices = { first, second };
			}

		/// The column just added holds one of these words, by value 0, 1, 2 ...
		void WordList (Table &t, std::vector<std::wstring> words)
			{
			t.bindings.back ().wordList = words;
			t.schema.cols.back ().choices = words;
			}

		const std::vector<std::wstring> kFeedTypes = { L"per rev", L"per min", L"surface finish" };

		/// A signed double as two columns: its size (never negative) and its sign.
		void SignedDouble (Table &t, const wchar_t *size, const wchar_t *sign, Acc at,
						   const wchar_t *negative, const wchar_t *positive)
			{
			AddOp (t, size, Kind::DoubleSize, at, false, 0.0);
			AddOp (t, sign, Kind::DoubleSign, at);
			Words (t, negative, positive);
			}

		/// FEEDS AND SPEEDS, HOME, REFERENCE POINTS AND PLANES - the same struct
		/// members for every kind of operation, so declared once.
		void AddOpLevel (Table &t)
			{
			// ---- Feeds and speeds. They live on the operation's tool info, not
			// in the toolpath struct, and the SIGN CARRIES MEANING: a feed's sign
			// is its unit (positive per minute, negative per revolution - plunge
			// and retract follow feed, as real parts show), and the speed's sign
			// is the spindle direction (positive CW, negative CCW). Each is TWO
			// columns: the size, always positive, and the sign as a word - so a
			// person never edits a sign by accident. The size is declared first:
			// a load applies changes in table order, and writing a size keeps
			// whatever sign is there, then the word sets it.
			SignedDouble (t, L"feed",    L"feed_mode",    ACCO (tl.feed),    L"per rev", L"per min");
			SignedDouble (t, L"plunge",  L"plunge_mode",  ACCO (tl.plunge),  L"per rev", L"per min");
			SignedDouble (t, L"retract", L"retract_mode", ACCO (tl.retract), L"per rev", L"per min");

			// speed is SURFACE speed when speed_mode is CSS (use_css), RPM otherwise.
			AddOp (t, L"speed",            Kind::LongSize, ACCO (tl.rpm), false, 0.0);
			AddOp (t, L"speed_mode",       Kind::BoolWord, ACCO (tl.use_css));
			Words (t, L"RPM", L"CSS");
			AddOp (t, L"spindle_dir",      Kind::LongSign, ACCO (tl.rpm));
			Words (t, L"CCW", L"CW");
			AddOp (t, L"max_ss",           Kind::Long,   ACCO (tl.max_ss), false, 0.0);
			AddOp (t, L"surf_fin_feed",    Kind::Bool,   ACCO (tl.surf_fin_feed));
			AddOp (t, L"plunge_surf_fin",  Kind::Bool,   ACCO (tl.pl_surf_fin_feed));

			// V9 coolant: one setting, Off / Flood / Mist / Thru-tool (a bit field on the
			// tool info). Only machines set up for V9 coolant use it - see Coolant.h.
			AddOp (t, L"coolant",          Kind::CoolantV9, ACCO (tl.coolant));

			// X-style coolant, by the machine's own names: which coolants come on
			// before, with and after the move. A write replaces only that timing's
			// coolant entries in the canned text - see Coolant.h.
			AddOp (t, L"coolant_before",   Kind::Coolant, ACCO (cantxt));
			t.bindings.back ().arg = Coolant::Before;
			AddOp (t, L"coolant_with",     Kind::Coolant, ACCO (cantxt));
			t.bindings.back ().arg = Coolant::With;
			AddOp (t, L"coolant_after",    Kind::Coolant, ACCO (cantxt));
			t.bindings.back ().arg = Coolant::After;

			// ---- Home position. always_on: 0 from machine, 1 user defined, 2 from tool.
			AddOp (t, L"home_mode",        Kind::Byte,   ACCO (home.always_on), false, 0, 2);
			AddOp (t, L"home_x",           Kind::Double, ACCO (home.pt[0]));
			AddOp (t, L"home_y",           Kind::Double, ACCO (home.pt[1]));
			AddOp (t, L"home_z",           Kind::Double, ACCO (home.pt[2]));

			// ---- Reference points: whether each is in use, where, and in which
			// frame. Type: 0 absolute world, 1 incremental world, 2 absolute tool
			// plane, 3 incremental tool plane.
			AddOp (t, L"ref_pt_on",        Kind::Bool,   ACCO (cmn.ref_pt_on));
			AddOp (t, L"ref_pt_x",         Kind::Double, ACCO (cmn.ref_pt[0]));
			AddOp (t, L"ref_pt_y",         Kind::Double, ACCO (cmn.ref_pt[1]));
			AddOp (t, L"ref_pt_z",         Kind::Double, ACCO (cmn.ref_pt[2]));
			AddOp (t, L"ref_pt2_on",       Kind::Bool,   ACCO (cmn.ref_pt2_on));
			AddOp (t, L"ref_pt2_x",        Kind::Double, ACCO (cmn.ref_pt2[0]));
			AddOp (t, L"ref_pt2_y",        Kind::Double, ACCO (cmn.ref_pt2[1]));
			AddOp (t, L"ref_pt2_z",        Kind::Double, ACCO (cmn.ref_pt2[2]));
			AddOp (t, L"ref_pt_state",     Kind::Byte,   ACCO (cmn.ref_pt_state));
			AddOp (t, L"approach_ref_type", Kind::Int,   ACCO (cmn.m_ApproachReferencePointType),
				   false, 0, 3);
			AddOp (t, L"retract_ref_type", Kind::Int,    ACCO (cmn.m_RetractReferencePointType),
				   false, 0, 3);

			// ---- The lathe entry point and retract point, and their clearances.
			AddOp (t, L"entry_pt_on",      Kind::Bool,   ACCO (cmn_lathe.entry_pt_on));
			AddOp (t, L"entry_pt_x",       Kind::Double, ACCO (cmn_lathe.entry_pt[0]));
			AddOp (t, L"entry_pt_y",       Kind::Double, ACCO (cmn_lathe.entry_pt[1]));
			AddOp (t, L"entry_pt_z",       Kind::Double, ACCO (cmn_lathe.entry_pt[2]));
			AddOp (t, L"retr_pt_on",       Kind::Bool,   ACCO (cmn_lathe.retr_pt_on));
			AddOp (t, L"clr_from_op",      Kind::Bool,   ACCO (cmn_lathe.clr_from_op));
			AddOp (t, L"min_ltool_clr",    Kind::Double, ACCO (cmn_lathe.min_ltool_clr), false, 0.0);
			AddOp (t, L"min_ee_clr",       Kind::Double, ACCO (cmn_lathe.min_ee_clr), false, 0.0);
			AddOp (t, L"retract_clearance", Kind::Double, ACCO (cmn_lathe.retract_clearance));
			AddOp (t, L"suppress_retract", Kind::Bool,   ACCO (cmn_lathe.suppressRetract));

			// ---- The planes: clearance, retract, feed, depth, top of stock. Each
			// is on/off, absolute or incremental, and a value.
			AddOp (t, L"clearance_on",     Kind::Bool,   ACCO (cmn.clearance_on));
			AddOp (t, L"clearance_inc",    Kind::Bool,   ACCO (cmn.clearance_inc));
			AddOp (t, L"clearance_pln",    Kind::Double, ACCO (cmn.clearance_pln));
			AddOp (t, L"clr_start_end",    Kind::Bool,   ACCO (cmn.clr_start_end));
			AddOp (t, L"retract_on",       Kind::Bool,   ACCO (cmn.retract_on));
			AddOp (t, L"retract_inc",      Kind::Bool,   ACCO (cmn.retract_inc));
			AddOp (t, L"retract_pln",      Kind::Double, ACCO (cmn.retract_pln));
			AddOp (t, L"feed_inc",         Kind::Bool,   ACCO (cmn.feed_inc));
			AddOp (t, L"feed_pln",         Kind::Double, ACCO (cmn.feed_pln));
			AddOp (t, L"depth_inc",        Kind::Bool,   ACCO (cmn.depth_inc));
			AddOp (t, L"depth",            Kind::Double, ACCO (cmn.depth));
			AddOp (t, L"top_stock_inc",    Kind::Bool,   ACCO (cmn.top_stock_inc));
			AddOp (t, L"top_stock",        Kind::Double, ACCO (cmn.top_stock));
			AddOp (t, L"stk_remain",       Kind::Double, ACCO (cmn.stk_remain));

			// ---- The toolpath filter (arc filter / tolerance): turns runs of short
			// moves into arcs and longer lines. reduce_type: 0 = neither, 1 =
			// reduce line tolerance, 2 = reduce arc tolerance.
			AddOp (t, L"filter_on",        Kind::Bool,   ACCO (filter.on));
			AddOp (t, L"filter_tol",       Kind::Double, ACCO (filter.max_error), false, 0.0);
			AddOp (t, L"filter_look_ahead", Kind::Short, ACCO (filter.look_ahead), false, 0, 10000);
			AddOp (t, L"filter_one_way",   Kind::Bool,   ACCO (filter.one_way));
			AddOp (t, L"filter_arcs_xy",   Kind::Bool,   ACCO (filter.create_arcs_xy));
			AddOp (t, L"filter_arcs_xz",   Kind::Bool,   ACCO (filter.create_arcs_xz));
			AddOp (t, L"filter_arcs_yz",   Kind::Bool,   ACCO (filter.create_arcs_yz));
			AddOp (t, L"filter_min_rad",   Kind::Double, ACCO (filter.min_radius), false, 0.0);
			AddOp (t, L"filter_max_rad",   Kind::Double, ACCO (filter.max_radius), false, 0.0);
			AddOp (t, L"filter_reduce_type", Kind::Byte, ACCO (filter.reduce_tol_type), false, 0, 2);
			AddOp (t, L"filter_reduced_tol", Kind::Double, ACCO (filter.reduced_tol_value), false, 0.0);

			// ---- WHERE THE OPERATION SITS: the tool plane and work coordinate
			// system it was made in. Shown, never written - the numbers identify
			// a named plane, and changing one from a spreadsheet would move the
			// whole toolpath into another coordinate frame.
			AddOp (t, L"tplane_id",        Kind::Short,  ACCO (tpln.m_PlaneID), true);
			AddOp (t, L"tplane_woff",      Kind::Short,  ACCO (tpln.woff_n), true);
			AddOp (t, L"wcs_id",           Kind::Short,  ACCO (WCS.m_PlaneID), true);
			AddOp (t, L"wcs_woff",         Kind::Short,  ACCO (WCS.woff_n), true);
			AddOp (t, L"group_id",         Kind::Long,   ACCO (cmn.grp_idn), true);
			}

		/// The tool inspection block. The same struct sits in all three
		/// parameter types, so it is declared once.
		/// The tool inspection columns, for a struct holding prm_ltool_inspect at PATH.
		#define ADD_INSPECT(S, PATH) \
			Add (t, L"insp_do_stop",       Kind::Bool,  ACC (S, PATH.do_stop)); \
			Add (t, L"insp_retract_type",  Kind::Byte,  ACC (S, PATH.stop_retract_type), \
				 false, 0, 1); \
			Add (t, L"insp_pt_x",          Kind::Double, ACC (S, PATH.stop_pt[0])); \
			Add (t, L"insp_pt_y",          Kind::Double, ACC (S, PATH.stop_pt[1])); \
			Add (t, L"insp_pt_z",          Kind::Double, ACC (S, PATH.stop_pt[2])); \
			Add (t, L"insp_move_x",        Kind::Bool,  ACC (S, PATH.stop_do_xyz[0])); \
			Add (t, L"insp_move_y",        Kind::Bool,  ACC (S, PATH.stop_do_xyz[1])); \
			Add (t, L"insp_move_z",        Kind::Bool,  ACC (S, PATH.stop_do_xyz[2])); \
			Add (t, L"insp_use_ref_pts",   Kind::Bool,  ACC (S, PATH.stop_use_ref_pts)); \
			Add (t, L"insp_abs_retract",   Kind::Bool,  ACC (S, PATH.stop_abs_retract)); \
			Add (t, L"insp_each_groove",   Kind::Bool,  ACC (S, PATH.stop_after_each_groove)); \
			Add (t, L"insp_each_depth",    Kind::Bool,  ACC (S, PATH.stop_after_each_depth)); \
			Add (t, L"insp_first_cut",     Kind::Bool,  ACC (S, PATH.stop_after_first_cut)); \
			Add (t, L"insp_n_cuts_on",     Kind::Bool,  ACC (S, PATH.stop_after_number_of_cuts)); \
			Add (t, L"insp_n_cuts",        Kind::Int,   ACC (S, PATH.stop_cuts), \
				 false, 0, 9999); \
			Add (t, L"insp_time_on",       Kind::Bool,  ACC (S, PATH.stop_after_time)); \
			Add (t, L"insp_time",          Kind::MinSec, ACC (S, PATH.stop_time)); \
			Add (t, L"insp_dist_on",       Kind::Bool,  ACC (S, PATH.stop_after_distance)); \
			Add (t, L"insp_dist",          Kind::Double, ACC (S, PATH.stop_distance), \
				 false, 0.0); \
			Add (t, L"insp_comment_on",    Kind::Bool,  ACC (S, PATH.stop_do_comment)); \
			Add (t, L"insp_comment",       Kind::Text,  ACC (S, PATH.stop_comment), \
				 false, kNone, kNone, \
				 sizeof (static_cast<S *> (nullptr)->PATH.stop_comment) / sizeof (TCHAR)); \
			Add (t, L"insp_between_cuts",  Kind::Bool,  ACC (S, PATH.stop_between_cuts)); \
			Add (t, L"insp_lead_dist",     Kind::Double, ACC (S, PATH.lead_distance), \
				 false, 0.0); \
			Add (t, L"insp_min_cut",       Kind::Double, ACC (S, PATH.min_cut), \
				 false, 0.0); \
			Add (t, L"insp_use_lead_in_out", Kind::Bool, ACC (S, PATH.use_lead_in_out)); \
			Add (t, L"insp_each_section",  Kind::Bool,  ACC (S, PATH.stopAfterEachSection)); \
			Add (t, L"insp_sections",      Kind::Short, ACC (S, PATH.numberOfSections), \
				 false, 0, 999); \
			Add (t, L"insp_at_end",        Kind::Bool,  ACC (S, PATH.stopAfterOperation));

		template <class S>
		void AddInspect (Table &t)
			{
			ADD_INSPECT (S, inspect);
			}

		/// Prime turning: its roughing passes' inspection (where the stops are).
		template <class S>
		void AddInspectPrime (Table &t)
			{
			ADD_INSPECT (S, rough.inspect);
			}

		Table MakeRough ()
			{
			using S = prm_lrough;
			Table t;
			t.opcode = TP_LROUGH;
			t.schema.type = L"ROUGH";

			// direction is shown and never written: it decides which way the
			// whole toolpath runs, and changing it from a spreadsheet is not
			// the same thing as changing a stepover.
			Add (t, L"direction",       Kind::Short,  ACC (S, direction), true);

			// ---- The stepover / depth of cut set.
			Add (t, L"step",            Kind::Double, ACC (S, step), false, 0.0);
			Add (t, L"equal_steps",     Kind::Bool,   ACC (S, equal_steps));
			Add (t, L"min_step",        Kind::Double, ACC (S, min_step), false, 0.0);
			Add (t, L"overlap",         Kind::Double, ACC (S, overlap), false, 0.0);
			Add (t, L"use_overlap",     Kind::Bool,   ACC (S, use_overlap));
			Add (t, L"doPeckDepths",    Kind::Bool,   ACC (S, doPeckDepths));
			Add (t, L"firstDepth",      Kind::Double, ACC (S, firstDepth), false, 0.0);
			Add (t, L"lastDepth",       Kind::Double, ACC (S, lastDepth), false, 0.0);
			Add (t, L"useDepthIncrement", Kind::Bool, ACC (S, useDepthIncrement));
			Add (t, L"depthIncrement",  Kind::Double, ACC (S, depthIncrement), false, 0.0);

			// ---- Stock, entry, exit.
			Add (t, L"stock_x",         Kind::Double, ACC (S, stock_x));
			Add (t, L"stock_z",         Kind::Double, ACC (S, stock_z));
			Add (t, L"entry_amt",       Kind::Double, ACC (S, entry_amt));
			Add (t, L"exit_amt",        Kind::Double, ACC (S, exit_amt));
			Add (t, L"rough_angle",     Kind::Double, ACC (S, rough_angle));
			Add (t, L"zigzag",          Kind::Byte,   ACC (S, zigzag), false, 0, 1);
			Add (t, L"rough_ltol",      Kind::Double, ACC (S, rough_ltol), false, 0.0);
			Add (t, L"remaining_stock", Kind::Bool,   ACC (S, remaining_stock));
			Add (t, L"shorten",         Kind::Bool,   ACC (S, shorten));

			// ---- Semi-finish passes.
			Add (t, L"do_finish",       Kind::Bool,   ACC (S, do_finish));
			Add (t, L"fin_n_cuts",      Kind::Short,  ACC (S, fin_n_cuts), false, 0, 999);
			Add (t, L"fin_step",        Kind::Double, ACC (S, fin_step), false, 0.0);
			Add (t, L"fin_stock_x",     Kind::Double, ACC (S, fin_stock_x));
			Add (t, L"fin_stock_z",     Kind::Double, ACC (S, fin_stock_z));

			// ---- Section turning.
			Add (t, L"doSectionTurning", Kind::Bool,  ACC (S, doSectionTurning));
			Add (t, L"numberOfSections", Kind::Short, ACC (S, numberOfSections),
				 false, 0, 999);

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			AddInspect<S> (t);
			return t;
			}

		Table MakeFinish ()
			{
			using S = prm_lfinish;
			Table t;
			t.opcode = TP_LFINISH;
			t.schema.type = L"FINISH";

			Add (t, L"direction",       Kind::Short,  ACC (S, direction), true);
			Add (t, L"n_cuts",          Kind::Short,  ACC (S, n_cuts), false, 0, 999);
			Add (t, L"step",            Kind::Double, ACC (S, step), false, 0.0);
			Add (t, L"stock_x",         Kind::Double, ACC (S, stock_x));
			Add (t, L"stock_z",         Kind::Double, ACC (S, stock_z));
			Add (t, L"fin_ltol",        Kind::Double, ACC (S, fin_ltol), false, 0.0);
			Add (t, L"use_backofs",     Kind::Bool,   ACC (S, use_backofs));
			Add (t, L"backofs_num",     Kind::Long,   ACC (S, backofs_num), false, 0, 9999);

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			AddInspect<S> (t);
			return t;
			}

		Table MakeDynamic ()
			{
			using S = prm_ldynamic;
			Table t;
			t.opcode = TP_LATHE_DYNAMIC_ROUGH;
			t.schema.type = L"DYNAMIC";

			Add (t, L"direction",       Kind::Short,  ACC (S, direction), true);
			Add (t, L"zigzag",          Kind::Bool,   ACC (S, zigzag));

			// Both an amount and a percent of tool radius are stored; the
			// dialog uses one, and which is in force is not something this
			// table knows. Edit the one you use.
			Add (t, L"stepover",        Kind::Double, ACC (S, stepover), false, 0.0);
			Add (t, L"stepover_percent", Kind::Double, ACC (S, stepover_percent), false, 0.0);
			Add (t, L"radius",          Kind::Double, ACC (S, radius), false, 0.0);
			Add (t, L"radius_percent",  Kind::Double, ACC (S, radius_percent), false, 0.0);

			Add (t, L"stock_x",         Kind::Double, ACC (S, stock_x));
			Add (t, L"stock_z",         Kind::Double, ACC (S, stock_z));
			Add (t, L"linear_tol",      Kind::Double, ACC (S, linear_tol), false, 0.0);
			Add (t, L"remaining_stock", Kind::Bool,   ACC (S, remaining_stock));
			Add (t, L"preventUpCutting", Kind::Bool,  ACC (S, preventUpCutting));
			Add (t, L"nonCuttingRegionAngle", Kind::Double, ACC (S, nonCuttingRegionAngle));

			Add (t, L"do_finish",       Kind::Bool,   ACC (S, do_finish));
			Add (t, L"fin_n_cuts",      Kind::Short,  ACC (S, fin_n_cuts), false, 0, 999);
			Add (t, L"fin_step",        Kind::Double, ACC (S, fin_step), false, 0.0);
			Add (t, L"fin_stock_x",     Kind::Double, ACC (S, fin_stock_x));
			Add (t, L"fin_stock_z",     Kind::Double, ACC (S, fin_stock_z));

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			AddInspect<S> (t);
			return t;
			}

		/// Lathe drill. No tool inspection block, so inspectStart is the end.
		Table MakeDrill ()
			{
			using S = prm_ldrill;
			Table t;
			t.opcode = TP_LDRILL;
			t.schema.type = L"DRILL";

			Add (t, L"cycle",           Kind::Short,  ACC (S, cycle), false, 0, 99);
			Add (t, L"peck1",           Kind::Double, ACC (S, peck1), false, 0.0);
			Add (t, L"peck2",           Kind::Double, ACC (S, peck2), false, 0.0);
			Add (t, L"peck_clr",        Kind::Double, ACC (S, peck_clr), false, 0.0);
			Add (t, L"chip_break",      Kind::Double, ACC (S, chip_break), false, 0.0);
			Add (t, L"dwell",           Kind::Double, ACC (S, dwell), false, 0.0);
			Add (t, L"shift",           Kind::Double, ACC (S, shift));
			Add (t, L"brk_thru",        Kind::Double, ACC (S, brk_thru));
			Add (t, L"drill_tip",       Kind::Bool,   ACC (S, drill_tip));
			Add (t, L"ldrillZ",         Kind::Double, ACC (S, ldrillZ));
			Add (t, L"ldrillX",         Kind::Double, ACC (S, ldrillX));
			Add (t, L"do_custom",       Kind::Bool,   ACC (S, do_custom));

			// Custom drill parameters 1 - 10: the post's drl_prm1$ - drl_prm10$.
			// What each means is the post's business (a two-touch probe,
			// custom cycle 18, reads 1 and 2). The SDK sends 0s unless do_custom.
			Add (t, L"custom1",         Kind::Double, ACC (S, custom[0]));
			Add (t, L"custom2",         Kind::Double, ACC (S, custom[1]));
			Add (t, L"custom3",         Kind::Double, ACC (S, custom[2]));
			Add (t, L"custom4",         Kind::Double, ACC (S, custom[3]));
			Add (t, L"custom5",         Kind::Double, ACC (S, custom[4]));
			Add (t, L"custom6",         Kind::Double, ACC (S, custom[5]));
			Add (t, L"custom7",         Kind::Double, ACC (S, custom[6]));
			Add (t, L"custom8",         Kind::Double, ACC (S, custom[7]));
			Add (t, L"custom9",         Kind::Double, ACC (S, custom[8]));
			Add (t, L"custom10",        Kind::Double, ACC (S, custom[9]));

			Add (t, L"clr_from_stk",    Kind::Bool,   ACC (S, clr_from_stk));
			Add (t, L"retr_from_stk",   Kind::Bool,   ACC (S, retr_from_stk));

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			return t;
			}

		/// Manual entry: the text that goes into the program. It has no tool, so
		/// no operation-level set and no inspection - every column after these
		/// is blank on its rows.
		Table MakeManual ()
			{
			using S = prm_manual_entry;
			Table t;
			t.opcode = TP_LMANUAL_ENTRY;
			t.opcode2 = TP_MANUAL_ENTRY;
			t.schema.type = L"MANUAL";

			// source / save decide where the text lives; shown, never written.
			Add (t, L"manual_source",   Kind::Short,  ACC (S, source), true);
			Add (t, L"manual_save",     Kind::Short,  ACC (S, save), true);
			// gcode is how the entry is output. CONFIRMED against real entries (the
			// SDK's own labels say "sequence numbers" and are misleading): 1005 =
			// output as a COMMENT, 1006 = output as CODE. 1007 and 1026 exist ("same
			// line" variants) and have not been seen, so an edit may only set 1005 or
			// 1006; a row already holding another value is left alone.
			Add (t, L"manual_gcode",    Kind::Short,  ACC (S, gcode), false, 1005, 1006);
			Add (t, L"manual_text",     Kind::Text,   ACC (S, comment), false, kNone, kNone,
				 sizeof (static_cast<S *> (nullptr)->comment) / sizeof (TCHAR));
			// The SDK buffer would hold 3112, but the operation dialog stops at 3111;
			// hold a load to what the dialog itself allows.
			t.schema.cols.back ().hi = 3111.0;

			t.opLevelStart = t.schema.cols.size ();
			t.inspectStart = t.schema.cols.size ();
			return t;
			}
		}

	namespace
		{
		/// Lathe face.
		Table MakeFace ()
			{
			using S = prm_lathe_face;
			Table t;
			t.opcode = TP_LFACE;
			t.schema.type = L"FACE";

			Add (t, L"do_rough",        Kind::Bool,   ACC (S, do_rough));
			Add (t, L"rough_step",      Kind::Double, ACC (S, rough_step), false, 0.0);
			Add (t, L"do_finish",       Kind::Bool,   ACC (S, do_finish));
			Add (t, L"finish_step",     Kind::Double, ACC (S, finish_step), false, 0.0);
			Add (t, L"n_cuts",          Kind::Short,  ACC (S, n_finish), false, 0, 999);
			Add (t, L"stock_z",         Kind::Double, ACC (S, stock_z));
			Add (t, L"face_leadin",     Kind::Double, ACC (S, leadin));
			Add (t, L"face_retract",    Kind::Double, ACC (S, retract));
			Add (t, L"face_retract_rapid", Kind::Bool, ACC (S, retract_rapid));
			Add (t, L"face_overcut",    Kind::Double, ACC (S, overcut));
			Add (t, L"face_from_center", Kind::Bool,  ACC (S, from_z_axis));
			Add (t, L"use_finish_feed", Kind::Bool,   ACC (S, use_finish_feed));
			Add (t, L"finish_feed",     Kind::Double, ACC (S, finish_feed), false, 0.0);
			Add (t, L"use_finish_ss",   Kind::Bool,   ACC (S, use_finish_ss));
			Add (t, L"finish_ss",       Kind::Long,   ACC (S, finish_ss), false, 0.0);
			Add (t, L"finish_ss_css",   Kind::Bool,   ACC (S, finish_ss_css));

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			AddInspect<S> (t);
			return t;
			}

		/// Groove and plunge rough share one parameter block.
		Table MakeGrooveLike (long opcode, const wchar_t *type)
			{
			using S = prm_lgroove;
			Table t;
			t.opcode = opcode;
			t.schema.type = type;

			Add (t, L"do_rough",        Kind::Bool,   ACC (S, rgh.do_rough));
			Add (t, L"rough_step",      Kind::Double, ACC (S, rgh.step), false, 0.0);
			Add (t, L"rough_n_steps",   Kind::Short,  ACC (S, rgh.n_steps), false, 0, 9999);
			// 0 = number of steps, 1 = step amount, 2 = percent of tool width
			Add (t, L"rough_step_by",   Kind::Byte,   ACC (S, rgh.use_incr), false, 0, 2);
			Add (t, L"rough_step_percent", Kind::Double, ACC (S, rgh.step_percent), false, 0.0);
			Add (t, L"stock_x",         Kind::Double, ACC (S, rgh.stock_x));
			Add (t, L"stock_z",         Kind::Double, ACC (S, rgh.stock_z));
			Add (t, L"rough_backoff",   Kind::Double, ACC (S, rgh.backoff), false, 0.0);
			Add (t, L"do_finish",       Kind::Bool,   ACC (S, fin.do_finish));
			Add (t, L"n_cuts",          Kind::Short,  ACC (S, fin.n_cuts), false, 0, 999);
			Add (t, L"finish_step",     Kind::Double, ACC (S, fin.step), false, 0.0);
			Add (t, L"fin_stock_x",     Kind::Double, ACC (S, fin.stock_x));
			Add (t, L"fin_stock_z",     Kind::Double, ACC (S, fin.stock_z));
			Add (t, L"use_finish_feed", Kind::Bool,   ACC (S, fin.use_finish_feed));
			Add (t, L"finish_feed",     Kind::Double, ACC (S, fin.finish_feed), false, 0.0);
			Add (t, L"use_finish_ss",   Kind::Bool,   ACC (S, fin.use_finish_ss));
			Add (t, L"finish_ss",       Kind::Long,   ACC (S, fin.finish_ss), false, 0.0);
			Add (t, L"finish_ss_css",   Kind::Bool,   ACC (S, fin.finish_ss_css));
			Add (t, L"retract_rapid",   Kind::Bool,   ACC (S, retract_rapid));
			Add (t, L"retract_feedrate", Kind::Double, ACC (S, retract_feedrate), false, 0.0);

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			AddInspect<S> (t);
			return t;
			}

		/// Prime turning. ITS FEEDS AND SPEEDS ARE ITS OWN: rough and finish each
		/// carry a spindle speed and axial / radial feeds inside this block, and
		/// those are what the toolpath uses - not the operation's feed and speed
		/// columns. Tool inspection is per pass type here and is not on the sheet.
		Table MakePrimeTurning ()
			{
			using S = PrmLathePrimeTurning;
			Table t;
			t.opcode = TP_LATHE_PRIME_TURNING;
			t.schema.type = L"PRIME";

			Add (t, L"direction",       Kind::Short,  ACC (S, common.direction), true);
			Add (t, L"linear_tol",      Kind::Double, ACC (S, common.linearTol), false, 0.0);
			Add (t, L"shorten",         Kind::Bool,   ACC (S, common.shortenPass));
			Add (t, L"remaining_stock", Kind::Bool,   ACC (S, common.remainingStock));

			// ---- Rough.
			Add (t, L"do_rough",        Kind::Bool,   ACC (S, rough.doRough));
			Add (t, L"pt_strategy",     Kind::ShortWord, ACC (S, rough.strategy));
			WordList (t, { L"horizontal", L"vertical", L"horizontal then vertical",
						   L"vertical then horizontal", L"alternate horizontal-vertical",
						   L"alternate vertical-horizontal" });
			Add (t, L"step",            Kind::Double, ACC (S, rough.cutDepth), false, 0.0);
			Add (t, L"min_step",        Kind::Double, ACC (S, rough.minCutDepth), false, 0.0);
			Add (t, L"equal_steps",     Kind::Bool,   ACC (S, rough.equalSteps));
			Add (t, L"stock_x",         Kind::Double, ACC (S, rough.stockX));
			Add (t, L"stock_z",         Kind::Double, ACC (S, rough.stockZ));
			Add (t, L"entry_amt",       Kind::Double, ACC (S, rough.entryAmount));
			Add (t, L"exit_amt",        Kind::Double, ACC (S, rough.exitAmount));
			Add (t, L"use_overlap",     Kind::Bool,   ACC (S, rough.doOverlap));
			Add (t, L"overlap",         Kind::Double, ACC (S, rough.overlap), false, 0.0);
			Add (t, L"rough_angle",     Kind::Double, ACC (S, rough.roughAngleRadians));

			Add (t, L"pt_rough_speed",  Kind::Long,   ACC (S, rough.feedSpeed.spindleSpeed), false, 0.0);
			Add (t, L"pt_rough_css",    Kind::Bool,   ACC (S, rough.feedSpeed.spindleSpeedCSS));
			Add (t, L"pt_rough_feed_from", Kind::ShortWord, ACC (S, rough.feedSpeed.feedFrom));
			WordList (t, { L"user defined", L"chip thickness" });
			Add (t, L"pt_rough_chip",   Kind::Double, ACC (S, rough.feedSpeed.chipThickness), false, 0.0);
			Add (t, L"pt_rough_max_feed", Kind::Double, ACC (S, rough.feedSpeed.maxFeedrate), false, 0.0);
			Add (t, L"pt_rough_feed_axial", Kind::Double, ACC (S, rough.feedSpeed.axialFeedrate), false, 0.0);
			Add (t, L"pt_rough_axial_type", Kind::ShortWord, ACC (S, rough.feedSpeed.axialFeedType));
			WordList (t, kFeedTypes);
			Add (t, L"pt_rough_feed_radial", Kind::Double, ACC (S, rough.feedSpeed.radialFeedrate), false, 0.0);
			Add (t, L"pt_rough_radial_type", Kind::ShortWord, ACC (S, rough.feedSpeed.radialFeedType));
			WordList (t, kFeedTypes);

			// ---- Finish.
			Add (t, L"do_finish",       Kind::Bool,   ACC (S, finish.doFinish));
			Add (t, L"finish_step",     Kind::Double, ACC (S, finish.cutDepth), false, 0.0);
			Add (t, L"n_cuts",          Kind::Short,  ACC (S, finish.numPasses), false, 0, 999);
			Add (t, L"fin_stock_x",     Kind::Double, ACC (S, finish.stockX));
			Add (t, L"fin_stock_z",     Kind::Double, ACC (S, finish.stockZ));

			Add (t, L"pt_fin_speed",    Kind::Long,   ACC (S, finish.feedSpeed.spindleSpeed), false, 0.0);
			Add (t, L"pt_fin_css",      Kind::Bool,   ACC (S, finish.feedSpeed.spindleSpeedCSS));
			Add (t, L"pt_fin_feed_from", Kind::ShortWord, ACC (S, finish.feedSpeed.feedFrom));
			WordList (t, { L"user defined", L"chip thickness" });
			Add (t, L"pt_fin_chip",     Kind::Double, ACC (S, finish.feedSpeed.chipThickness), false, 0.0);
			Add (t, L"pt_fin_max_feed", Kind::Double, ACC (S, finish.feedSpeed.maxFeedrate), false, 0.0);
			Add (t, L"pt_fin_feed_axial", Kind::Double, ACC (S, finish.feedSpeed.axialFeedrate), false, 0.0);
			Add (t, L"pt_fin_axial_type", Kind::ShortWord, ACC (S, finish.feedSpeed.axialFeedType));
			WordList (t, kFeedTypes);
			Add (t, L"pt_fin_feed_radial", Kind::Double, ACC (S, finish.feedSpeed.radialFeedrate), false, 0.0);
			Add (t, L"pt_fin_radial_type", Kind::ShortWord, ACC (S, finish.feedSpeed.radialFeedType));
			WordList (t, kFeedTypes);

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();	// no inspection columns for this kind
			AddInspectPrime<S> (t);
			return t;
			}
		}

	namespace
		{
		/// MILL depth cuts and multi passes - on the operation itself, shared by
		/// every mill kind.
		void AddMillCuts (Table &t)
			{
			AddOp (t, L"dcuts_on",       Kind::Bool,   ACCO (dcuts.on));
			AddOp (t, L"dcut_rough",     Kind::Double, ACCO (dcuts.rgh_amt), false, 0.0);
			AddOp (t, L"dcut_finish",    Kind::Double, ACCO (dcuts.fin_amt), false, 0.0);
			AddOp (t, L"dcut_fin_n",     Kind::Short,  ACCO (dcuts.fin_n), false, 0, 999);
			AddOp (t, L"dcut_stock",     Kind::Double, ACCO (dcuts.stock_t_l));
			AddOp (t, L"dcut_by_depth",  Kind::Bool,   ACCO (dcuts.by_depth));
			AddOp (t, L"mcuts_on",       Kind::Bool,   ACCO (mcuts.on));
			AddOp (t, L"mcut_rough_n",   Kind::Short,  ACCO (mcuts.rgh_n), false, 0, 999);
			AddOp (t, L"mcut_rough_amt", Kind::Double, ACCO (mcuts.rgh_amt), false, 0.0);
			AddOp (t, L"mcut_fin_n",     Kind::Short,  ACCO (mcuts.fin_n), false, 0, 999);
			AddOp (t, L"mcut_fin_amt",   Kind::Double, ACCO (mcuts.fin_amt), false, 0.0);
			}

		/// Mill contour (live tooling).
		Table MakeMillContour ()
			{
			using S = prm_contour;
			Table t;
			t.opcode = TP_CONTOUR;
			t.schema.type = L"CONTOUR";

			Add (t, L"taper",           Kind::Bool,   ACC (S, taper));
			Add (t, L"taper_ang",       Kind::Double, ACC (S, taper_ang));
			Add (t, L"in_corner_rad",   Kind::Double, ACC (S, in_corner_rad), false, 0.0);
			Add (t, L"ex_corner_rad",   Kind::Double, ACC (S, ex_corner_rad), false, 0.0);
			Add (t, L"fr_override_on",  Kind::Bool,   ACC (S, fr_override_on));
			Add (t, L"fr_override",     Kind::Double, ACC (S, fr_override), false, 0.0);
			Add (t, L"ss_override_on",  Kind::Bool,   ACC (S, ss_override_on));
			Add (t, L"ss_override",     Kind::Long,   ACC (S, ss_override), false, 0.0);
			AddMillCuts (t);

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			return t;
			}

		/// Mill drill: the same cycle settings as lathe drill.
		Table MakeMillDrill ()
			{
			using S = prm_drill;
			Table t;
			t.opcode = TP_DRILL;
			t.schema.type = L"MILL DRILL";

			Add (t, L"cycle",           Kind::Short,  ACC (S, cycle), false, 0, 99);
			Add (t, L"peck1",           Kind::Double, ACC (S, peck1), false, 0.0);
			Add (t, L"peck2",           Kind::Double, ACC (S, peck2), false, 0.0);
			Add (t, L"peck_clr",        Kind::Double, ACC (S, peck_clr), false, 0.0);
			Add (t, L"chip_break",      Kind::Double, ACC (S, chip_break), false, 0.0);
			Add (t, L"dwell",           Kind::Double, ACC (S, dwell), false, 0.0);
			Add (t, L"shift",           Kind::Double, ACC (S, shift));
			Add (t, L"brk_thru",        Kind::Double, ACC (S, brk_thru));
			Add (t, L"drill_tip",       Kind::Bool,   ACC (S, drill_tip));
			Add (t, L"do_custom",       Kind::Bool,   ACC (S, do_custom));

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			return t;
			}

		/// 2D high-speed: dynamic mill, peel mill, core / area / blend / rest.
		/// Its back feed is the feed of the repositioning moves - on the sheet so
		/// the time estimate can follow it.
		Table MakeDynamicMill ()
			{
			using S = prm_2d_hmm;
			Table t;
			t.opcode = TP_2D_HMM;
			t.schema.type = L"DYNAMIC MILL";

			Add (t, L"hmm_style",       Kind::Byte,   ACC (S, style), true);
			Add (t, L"stepover",        Kind::Double, ACC (S, step), false, 0.0);
			Add (t, L"back_feed",       Kind::Double, ACC (S, back_feed), false, 0.0);
			Add (t, L"hmm_finish_pass", Kind::Bool,   ACC (S, fin_pass));
			Add (t, L"hmm_ext_entry",   Kind::Bool,   ACC (S, ext_entry));
			Add (t, L"hmm_entry_dist",  Kind::Double, ACC (S, entry_dist), false, 0.0);
			Add (t, L"hmm_ext_exit",    Kind::Bool,   ACC (S, ext_exit));
			Add (t, L"hmm_exit_dist",   Kind::Double, ACC (S, exit_dist), false, 0.0);
			AddMillCuts (t);

			t.opLevelStart = t.schema.cols.size ();
			AddOpLevel (t);
			t.inspectStart = t.schema.cols.size ();
			return t;
			}
		}

	const std::vector<Table> &AllTables ()
		{
		static const std::vector<Table> tables = { MakeRough (), MakeFinish (),
												   MakeDynamic (), MakeDrill (),
												   MakeManual (), MakeFace (), MakePrimeTurning (),
											   MakeGrooveLike (TP_LPLUNGE_ROUGH, L"PLUNGE ROUGH"),
											   MakeGrooveLike (TP_LGROOVE, L"GROOVE"),
											   MakeMillContour (), MakeMillDrill (),
											   MakeDynamicMill () };
		return tables;
		}

	const std::vector<std::wstring> &SheetColumns ()
		{
		static const std::vector<std::wstring> cols = [] ()
			{
			std::vector<std::wstring> out;
			auto have = [&out] (const std::wstring &n)
				{
				return std::find (out.begin (), out.end (), n) != out.end ();
				};

			const std::vector<Table> &all = AllTables ();

			// The toolpath-specific columns of each kind, in table order. A name
			// two kinds share (stock_x, step, direction) is ONE column: the row's
			// type says which struct it reads from.
			for (const Table &t : all)
				for (size_t i = 0; i < t.opLevelStart; ++i)
					if (!have (t.schema.cols[i].name))
						out.push_back (t.schema.cols[i].name);

			// The operation-level and inspection sets are the same in every
			// table, so the first one speaks for all.
			const Table &first = all.front ();
			for (size_t i = first.opLevelStart; i < first.schema.cols.size (); ++i)
				if (!have (first.schema.cols[i].name))
					out.push_back (first.schema.cols[i].name);
			return out;
			} ();
		return cols;
		}

	int IndexOf (const Table &t, const std::wstring &name)
		{
		for (size_t i = 0; i < t.schema.cols.size (); ++i)
			if (t.schema.cols[i].name == name)
				return static_cast<int> (i);
		return -1;
		}

	const Table *TableFor (long opcode)
		{
		for (const Table &t : AllTables ())
			if (t.opcode == opcode || (t.opcode2 != 0 && t.opcode2 == opcode))
				return &t;
		return nullptr;
		}

	void *PrmFor (operation &op, long opcode)
		{
		switch (opcode)
			{
			case TP_LROUGH:               return &op.u.prm_lrgh;
			case TP_LFINISH:              return &op.u.prm_lfin;
			case TP_LATHE_DYNAMIC_ROUGH:  return &op.u.prm_ldynamic;
			case TP_LDRILL:               return &op.u.prm_ldrl;
			case TP_LMANUAL_ENTRY:        return &op.u.manual_entry;
			case TP_LFACE:                return &op.u.prm_lface;
			case TP_LATHE_PRIME_TURNING:  return &op.u.prmLathePrimeTurning;
			case TP_LPLUNGE_ROUGH:
			case TP_LGROOVE:              return &op.u.prm_lgrv;
			case TP_CONTOUR:              return &op.u.prm_cntr;
			case TP_DRILL:                return &op.u.prm_drl;
			case TP_2D_HMM:               return &op.u.hmm_2d;
			}
		return nullptr;
		}

	std::wstring Read (const Binding &b, void *op, void *prm)
		{
		void *p = b.at (b.base == Base::Op ? op : prm);
		switch (b.kind)
			{
			case Kind::Bool:
				return *static_cast<bool *> (p) ? L"1" : L"0";
			case Kind::Byte:
				return std::to_wstring (static_cast<int> (*static_cast<MC_BYTE *> (p)));
			case Kind::Short:
				return std::to_wstring (static_cast<int> (*static_cast<short *> (p)));
			case Kind::Int:
				return std::to_wstring (*static_cast<int *> (p));
			case Kind::Long:
				return std::to_wstring (*static_cast<long *> (p));
			case Kind::Double:
				return Csv::FormatDouble (*static_cast<double *> (p));
			case Kind::Coolant:
				return Coolant::Describe (*static_cast<const operation *> (op), b.arg);
			case Kind::CoolantV9:
				return Coolant::DescribeV9 (*static_cast<short *> (p));
			case Kind::MinSec:
				return Inspect::MinSec (*static_cast<double *> (p));
			case Kind::DoubleSize:
				return Csv::FormatDouble (std::fabs (*static_cast<double *> (p)));
			case Kind::DoubleSign:
				return b.words[std::signbit (*static_cast<double *> (p)) ? 0 : 1];
			case Kind::LongSize:
				return std::to_wstring (std::labs (*static_cast<long *> (p)));
			case Kind::LongSign:
				return b.words[*static_cast<long *> (p) < 0 ? 0 : 1];
			case Kind::BoolWord:
				return b.words[*static_cast<bool *> (p) ? 1 : 0];
			case Kind::ShortWord:
				{
				const short v = *static_cast<short *> (p);
				if (v >= 0 && static_cast<size_t> (v) < b.wordList.size ())
					return b.wordList[static_cast<size_t> (v)];
				return std::to_wstring (v);		// a value with no word: shown, not editable to
				}
			case Kind::Text:
				{
				const TCHAR *s = static_cast<const TCHAR *> (p);
				size_t n = 0;
				while (n < b.textLen && s[n] != 0)
					++n;
				return std::wstring (s, n);
				}
			}
		return L"";
		}

	bool Write (const Binding &b, void *op, void *prm, const std::wstring &text)
		{
		void *p = b.at (b.base == Base::Op ? op : prm);
		switch (b.kind)
			{
			case Kind::Bool:
				{
				bool v = false;
				if (!Csv::ParseBool (text, v))
					return false;
				*static_cast<bool *> (p) = v;
				return true;
				}
			case Kind::Byte:
			case Kind::Short:
			case Kind::Int:
			case Kind::Long:
				{
				long long v = 0;
				if (!Csv::ParseLong (text, v))
					return false;
				if (b.kind == Kind::Byte)
					{
					if (v < 0 || v > 255)
						return false;
					*static_cast<MC_BYTE *> (p) = static_cast<MC_BYTE> (v);
					}
				else if (b.kind == Kind::Short)
					{
					if (v < -32768 || v > 32767)
						return false;
					*static_cast<short *> (p) = static_cast<short> (v);
					}
				else if (b.kind == Kind::Int)
					*static_cast<int *> (p) = static_cast<int> (v);
				else
					*static_cast<long *> (p) = static_cast<long> (v);
				return true;
				}
			case Kind::Double:
				{
				double v = 0;
				if (!Csv::ParseDouble (text, v))
					return false;
				*static_cast<double *> (p) = v;
				return true;
				}
			case Kind::Coolant:
				{
				std::wstring why;
				return Coolant::Apply (*static_cast<operation *> (op), b.arg, text, why);
				}
			case Kind::CoolantV9:
				{
				std::wstring why;
				return Coolant::ApplyV9 (*static_cast<operation *> (op), text, why);
				}
			case Kind::MinSec:
				{
				double v = 0;
				if (!Inspect::ParseMinSec (text, v))
					return false;
				*static_cast<double *> (p) = v;
				return true;
				}
			case Kind::DoubleSize:
				{
				double v = 0;
				if (!Csv::ParseDouble (text, v) || v < 0)
					return false;
				double &d = *static_cast<double *> (p);
				d = std::copysign (v, d);			// the sign stays
				return true;
				}
			case Kind::LongSize:
				{
				long long v = 0;
				if (!Csv::ParseLong (text, v) || v < 0 || v > LONG_MAX)
					return false;
				long &l = *static_cast<long *> (p);
				l = l < 0 ? -static_cast<long> (v) : static_cast<long> (v);
				return true;
				}
			case Kind::ShortWord:
				{
				for (size_t i = 0; i < b.wordList.size (); ++i)
					if (_wcsicmp (Csv::Trim (text).c_str (), b.wordList[i].c_str ()) == 0)
						{
						*static_cast<short *> (p) = static_cast<short> (i);
						return true;
						}
				return false;
				}
			case Kind::DoubleSign:
			case Kind::LongSign:
			case Kind::BoolWord:
				{
				int which = -1;
				for (int i = 0; i < 2; ++i)
					if (b.words[i] != nullptr && _wcsicmp (Csv::Trim (text).c_str (), b.words[i]) == 0)
						which = i;
				if (which < 0)
					return false;
				if (b.kind == Kind::BoolWord)
					*static_cast<bool *> (p) = which == 1;
				else if (b.kind == Kind::DoubleSign)
					{
					double &d = *static_cast<double *> (p);
					d = std::copysign (std::fabs (d), which == 0 ? -1.0 : 1.0);
					}
				else
					{
					long &l = *static_cast<long *> (p);
					// A zero speed has no direction to hold - refuse rather than
					// drop the change silently.
					if (l == 0 && which == 0)
						return false;
					l = which == 0 ? -std::labs (l) : std::labs (l);
					}
				return true;
				}
			case Kind::Text:
				{
				if (text.size () + 1 > b.textLen)
					return false;
				TCHAR *s = static_cast<TCHAR *> (p);
				for (size_t i = 0; i < text.size (); ++i)
					s[i] = text[i];
				for (size_t i = text.size (); i < b.textLen; ++i)
					s[i] = 0;
				return true;
				}
			}
		return false;
		}
	}
