#include "ColumnHelp.h"

#include <map>

namespace ColumnHelp
	{
	std::wstring For (const std::wstring &name)
		{
		static const std::map<std::wstring, std::wstring> help = {
			// Identity and context
			{ L"op_idn",        L"The operation's ID - how a load finds its row's operation. Do not change." },
			{ L"type",          L"The kind of operation - ROUGH, FINISH, FACE, GROOVE, DRILL, MANUAL, DYNAMIC MILL and so on." },
			{ L"tool",          L"Tool number." },
			{ L"comment",       L"The operation's comment, as the Operation Manager shows it. Edit it here and it loads back (up to 119 characters)." },
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
			{ L"coolant",        L"V9 coolant - the one coolant setting on a machine set up for V9 coolant: Off, Flood, Mist or Thru-tool." },
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
			{ L"manual_text",    L"The manual entry's text (up to 3111 characters). Alt+Enter for a new line; with macros, double-click for a bigger editor." },
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

			// Tool inspection and insert flips
			{ L"insp_time",      L"Tool inspection: feed time between stops, as minutes:seconds (9:00) - or seconds (540)." },
			{ L"insp_dist",      L"Tool inspection: cut length between stops." },
			{ L"insp_between_cuts", L"1 = stop only between cuts (the real interval runs over the setting); 0 = stop mid-cut." },
			{ L"insp_min_cut",   L"Mid-cut stops: a pass with less than this left is finished first." },
			{ L"insp_at_end",    L"Tool inspection stop at the end of the operation." },
			{ L"flips",          L"Insert flips in the toolpath as it is: inspection stops whose comment changes the insert (CHANGE/ROTATE INSERT)." },
			{ L"flips_est",      L"Insert flips per part, LIVE: follows feeds, speeds and the inspection settings. Equal to flips until something changes." },
			{ L"flips_why",      L"Why each flip happened: time, distance, cuts (number of cuts, first cut, each depth) or end of the operation." },
			{ L"flip_longest",   L"The most feed time between flips (or from the start / to the end) - the longest any edge cuts." },
			{ L"insp_mode",      L"How stops are placed: between cuts only, or mid-cut (finishing a pass that is nearly done)." },
			{ L"cut_seconds_est", L"Feed time in seconds, LIVE - the cutting part of the estimate; the Tools page sums it per tool." },
			{ L"flips_uncommented", L"Flips from stops with NO comment, LIVE - inferred from the settings (they leave nothing in the NCI): one each time the edge has cut edge_limit, carrying short ops' time over to the tool's next op." },
			{ L"flips_part",     L"All the flips of the op per part, LIVE: commented (flips_est) + inferred (flips_uncommented). The Tools page sums it per tool." },
			{ L"edge_limit",     L"Cut time (seconds) an edge takes before a comment-less stop counts as a flip: the op's insp_time when filled in, else the tool's edge life on the Tools page (8:00)." },
			{ L"edge_after",     L"Cut time (seconds) the tool's edge has done since its last flip, leaving this op - the next op of the same tool carries on from it." },
			{ L"cut_dia",       L"Mean cutting diameter over the feed moves (from the NCI) - where an RPM op's surface speed, and CSS's max_ss cap, are taken. A drill: the tool diameter." },
			{ L"mrr",            L"Metal removal rate while cutting, LIVE, in3/min (cm3/min metric). Turning: 12 x SFM x feed/rev x depth; drill: pi D^2/4 x feed/min; mill: ae x ap x feed/min (see mrr_basis). Follows speed, feed and depth edits." },
			{ L"removed",        L"Material the op removes (in3, cm3 metric) - fixed by its toolpath, so feed and speed edits do not change it (a depth edit does only once the op is regenerated). From a 2D stock simulation where it covers the op, else mrr x cut time - see removed_from." },
			{ L"removed_from",   L"Where removed comes from: the stock simulation (the toolpath swept through the stock), or an estimate (mrr x cut time, which counts air as cutting). Drill cycles, mill moves and ops after a stock flip are not simulated." },
			{ L"xf_copies",      L"A transform: how many copies of its source operations its toolpath holds (measured). A source: how many more times transforms cut it - the Tools page counts its tool time and flips that many more times." },
			{ L"removed_check",  L"The stock simulation set against Mastercam's own stock boundaries: 'agrees', or both figures when they differ by more than 20% - usually a tool whose outline in Mastercam is not its real insert." },
			{ L"air_pct",        L"Share of the op's feed time cutting nothing (stock simulation)." },
			{ L"mrr_avg",        L"Actual MRR over the whole operation, LIVE: removed / op time - rapids, leads and air cuts included. Follows feed and speed edits through the estimate." },
			{ L"mrr_engaged",    L"Actual MRR while in metal, LIVE: removed / (cut time less its air share). Set it beside mrr (theoretical) - a big gap says the depth or engagement is not what the parameters suggest." },
			{ L"mrr_basis",      L"What mrr is worked from: the depth-of-cut column, the diameter, and whether max_ss caps the speed. A mill op: its ae (contour: full tool width unless multi passes - an upper bound; dynamic: stepover) and ap." },
			{ L"xf_type",        L"Transform kind: mirror, rotate or translate. Transforms are shown, never written." },
			{ L"xf_detail",      L"The transform in words: its kind and numbers (steps, angle or distance, about where)." },
			{ L"xf_instances",   L"Copies the transform makes, as its dialog counts them: rotate - its steps; rectangular array - X steps x Y steps; mirror - 1." },
			{ L"xf_sources",     L"The op_idn of each operation the transform copies, as Mastercam lists them." },
			{ L"xf_make_ops",    L"1 = the transform makes new operations (and geometry); 0 = it transforms the toolpath only." },
			{ L"xf_source_geom", L"1 = transforms the source operations' geometry; 0 = their NCI." },
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
