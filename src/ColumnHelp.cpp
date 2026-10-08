#include "ColumnHelp.h"

#include <map>

namespace ColumnHelp
	{
	std::wstring For (const std::wstring &name)
		{
		static const std::map<std::wstring, std::wstring> help = {
			// Identity and context
			{ L"op_idn",        L"The operation's ID - how a load finds its row's operation. Do not change." },
			{ L"type",          L"The kind of operation: ROUGH, FINISH, DYNAMIC, DRILL or MANUAL." },
			{ L"tool",          L"Tool number." },
			{ L"comment",       L"The operation's comment, as the Operation Manager shows it." },
			{ L"changes",       L"How many cells in this row differ from the dump. Edited cells are highlighted." },
			{ L"tool_radius",   L"The tool radius the percent columns are a percent of (dynamic rough)." },
			{ L"tool_name",     L"The tool's description." },
			{ L"group_name",    L"The toolpath group the operation is in." },
			{ L"units",         L"in or mm - the units every length on this row is in." },
			{ L"needs_regen",   L"yes = the toolpath is out of date; regenerate before posting. Stats are stale until then." },

			// Toolpath stats
			{ L"cycle_time",     L"Mastercam's estimated cycle time for this operation, h:mm:ss (matches the Operation Manager to within seconds)." },
			{ L"cycle_time_raw", L"The cycle time in seconds, as Mastercam returned it." },
			{ L"travel_x_min",   L"Smallest X the toolpath reaches - a RADIUS (double it for diameter)." },
			{ L"travel_x_max",   L"Largest X the toolpath reaches - a RADIUS (double it for diameter)." },
			{ L"travel_z_min",   L"Smallest Z the toolpath reaches, in the tool plane." },
			{ L"feed_groups",    L"Feed-move length at each feed rate, longest first - a dynamic mill's back feed is its own entry." },
			{ L"est_cycle_time", L"LIVE estimate: Mastercam's cycle time, moved by your feed / speed / CSS / max_ss edits on this row. Equals cycle_time until you edit." },
			{ L"time_change",    L"How much the live estimate differs from Mastercam's cycle time." },
			{ L"est_seconds",    L"The live estimate in seconds (the formula behind est_cycle_time)." },
			{ L"back_feed",      L"Dynamic mill back feed - the feed of the repositioning moves (per minute)." },
			{ L"dcut_rough",     L"Mill depth cuts: max rough step." },
			{ L"dcut_finish",    L"Mill depth cuts: finish step." },
			{ L"dcut_fin_n",     L"Mill depth cuts: number of finish cuts." },
			{ L"mcut_rough_n",   L"Mill multi passes: number of rough passes." },
			{ L"mcut_fin_n",     L"Mill multi passes: number of finish passes." },
			{ L"fr_override",    L"Contour feed rate override, used when fr_override_on is 1." },
			{ L"ss_override",    L"Contour spindle speed override, used when ss_override_on is 1." },
			{ L"cut_length",     L"Total length of the feed moves (lines and arcs) in the toolpath." },
			{ L"rapid_length",   L"Total length of the rapid moves WITHIN the operation - backplot also counts the moves to and from it, so its figure is larger." },
			{ L"travel_z_max",   L"Largest Z the toolpath reaches, in the tool plane." },

			// Feeds and speeds
			{ L"feed",           L"Cutting feed rate - always positive here. Per rev or per minute is set in feed_mode." },
			{ L"feed_mode",      L"per rev (IPR / mm per rev) or per min (IPM / mm per min) for feed." },
			{ L"speed",          L"Spindle speed - surface speed when speed_mode is CSS, RPM when it is RPM." },
			{ L"speed_mode",     L"CSS = constant surface speed (speed is SFM or m/min); RPM = fixed spindle speed." },
			{ L"spindle_dir",    L"CW (M03) or CCW (M04)." },
			{ L"max_ss",         L"Maximum spindle speed (RPM) - the cap when running CSS." },
			{ L"plunge",         L"Plunge feed rate - per rev or per minute is set in plunge_mode." },
			{ L"plunge_mode",    L"per rev or per min for plunge." },
			{ L"retract",        L"Retract feed rate - per rev or per minute is set in retract_mode." },
			{ L"retract_mode",   L"per rev or per min for retract." },
			{ L"surf_fin_feed",  L"1 = feed holds a surface finish value instead of a feed rate." },
			{ L"plunge_surf_fin", L"1 = plunge holds a surface finish value instead of a feed rate." },

			// Depth of cut
			{ L"step",             L"Rough: depth of cut per pass. Finish: stepover between finish passes." },
			{ L"min_step",         L"Smallest depth of cut the rough will take." },
			{ L"stepover",         L"Dynamic rough stepover. Change this and stepover_percent follows." },
			{ L"stepover_percent", L"Stepover as a percent of tool radius - calculated from stepover. Typing here replaces the formula." },
			{ L"n_cuts",           L"Number of finish passes." },
			{ L"peck1",            L"Drill: first peck depth." },
			{ L"peck2",            L"Drill: subsequent peck depth." },

			// Stock to leave
			{ L"stock_x",      L"Stock to leave in X." },
			{ L"stock_z",      L"Stock to leave in Z." },
			{ L"fin_stock_x",  L"Semi-finish passes: stock to leave in X." },
			{ L"fin_stock_z",  L"Semi-finish passes: stock to leave in Z." },

			// Coolant
			{ L"coolant_before", L"Coolant turned on BEFORE the move - pick from this machine's coolants, or none." },
			{ L"coolant_with",   L"Coolant turned on WITH the move - pick from this machine's coolants, or none." },
			{ L"coolant_after",  L"Coolant turned on AFTER the move - pick from this machine's coolants, or none." },
			{ L"coolant",        L"Old-style (V9) coolant bit field: 8 off, 16 flood, 32 mist, 64 through-tool." },
			{ L"coolant_text",   L"The V9 coolant field in words." },
			{ L"canned_text_raw", L"The operation's canned text and coolant codes, raw." },

			// Prime turning's own feeds and speeds
			{ L"pt_rough_speed",       L"Prime turning ROUGH spindle speed - CSS (SFM / m/min) when pt_rough_css is 1, else RPM." },
			{ L"pt_rough_css",         L"1 = prime turning rough speed is CSS." },
			{ L"pt_rough_feed_axial",  L"Prime turning ROUGH feed along Z; unit in pt_rough_axial_type." },
			{ L"pt_rough_feed_radial", L"Prime turning ROUGH feed along X; unit in pt_rough_radial_type." },
			{ L"pt_rough_feed_from",   L"user defined = the feeds as typed; chip thickness = Mastercam works them out from pt_rough_chip." },
			{ L"pt_fin_speed",         L"Prime turning FINISH spindle speed - CSS when pt_fin_css is 1, else RPM." },
			{ L"pt_fin_feed_axial",    L"Prime turning FINISH feed along Z; unit in pt_fin_axial_type." },
			{ L"pt_fin_feed_radial",   L"Prime turning FINISH feed along X; unit in pt_fin_radial_type." },
			{ L"pt_strategy",          L"Prime turning roughing strategy." },
			{ L"finish_feed",          L"Finish-pass feed rate, used when use_finish_feed is 1." },
			{ L"finish_ss",            L"Finish-pass spindle speed, used when use_finish_ss is 1." },
			{ L"rough_step",           L"Face: rough stepover. Groove / plunge rough: step across the groove." },
			{ L"finish_step",          L"Finish pass cut amount." },
			{ L"rough_step_by",        L"Groove rough step from: 0 number of steps, 1 step amount, 2 percent of tool width." },

			// Toolpath-specific
			{ L"direction",      L"Cut direction: 0 OD, 1 ID, 2 face, 3 back. Shown, never written." },
			{ L"overlap",        L"Rough: overlap amount between passes." },
			{ L"entry_amt",      L"Rough: entry amount." },
			{ L"exit_amt",       L"Rough: exit amount." },
			{ L"rough_angle",    L"Rough: cut angle, in RADIANS." },
			{ L"zigzag",         L"0 = one way, 1 = zigzag." },
			{ L"do_finish",      L"1 = take semi-finish passes." },
			{ L"fin_n_cuts",     L"Number of semi-finish passes." },
			{ L"fin_step",       L"Semi-finish stepover." },
			{ L"radius",         L"Dynamic rough toolpath radius. Change this and radius_percent follows." },
			{ L"radius_percent", L"Toolpath radius as a percent of tool radius - calculated from radius. Typing here replaces the formula." },
			{ L"linear_tol",     L"Tolerance for turning spline geometry into lines. Only matters on spline chains." },
			{ L"nonCuttingRegionAngle", L"Dynamic rough non-cutting region angle, degrees." },
			{ L"manual_gcode",   L"Manual entry output: 1005 = as a comment, 1006 = as code." },
			{ L"manual_text",    L"The manual entry's text (up to 3111 characters). Alt+Enter for a new line." },
			{ L"manual_source",  L"Where the manual entry text comes from. Shown, never written." },
			{ L"manual_save",    L"How the manual entry is saved. Shown, never written." },
			{ L"cycle",          L"Drill cycle number." },

			// Home, planes, filter
			{ L"home_mode",      L"Home position: 0 from machine, 1 user defined, 2 from tool." },
			{ L"clearance_pln",  L"Clearance plane value." },
			{ L"retract_pln",    L"Retract plane value." },
			{ L"feed_pln",       L"Feed plane value." },
			{ L"filter_on",      L"1 = arc filter / tolerance on: runs of short moves become arcs and longer lines." },
			{ L"filter_tol",     L"Arc filter tolerance." },
			{ L"filter_reduce_type", L"0 = neither, 1 = reduce line tolerance, 2 = reduce arc tolerance." },

			// Where it sits
			{ L"tplane_id",  L"Tool plane ID. Shown, never written." },
			{ L"wcs_id",     L"Work coordinate system ID. Shown, never written." },
			{ L"group_id",   L"Toolpath group ID. Shown, never written." },
			};
		const auto it = help.find (name);
		if (it != help.end ())
			return it->second;
		if (name.compare (0, 5, L"insp_") == 0)
			return L"Tool inspection setting.";
		if (name.compare (0, 7, L"filter_") == 0)
			return L"Arc filter / tolerance setting.";
		return std::wstring ();
		}
	}
