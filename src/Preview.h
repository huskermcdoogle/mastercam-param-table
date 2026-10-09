//
// Preview.h - the load's "here is exactly what will change" window.
//
// Every change a load would write, grouped by operation, old value struck
// through in red and new value in green; refused rows in amber with the
// reason; anything skipped or ignored listed too. Apply writes, Cancel does
// nothing - the same decision the plain Yes/No box used to ask, made where
// the whole change set can be read.
//
// EVERY CHANGE HAS A TICK BOX. Unticked, it is left out; an operation's own
// box ticks or unticks all of its changes. Changes that only make sense
// together (a feed and its per rev / per min) share a `link` and are ticked
// together - a feed of 0.3 applied without its "per rev" would be 0.3 per
// minute. The ticking rules are SDK-free (PreviewTicks.cpp) and tested.
//
#pragma once

#include <functional>
#include <string>
#include <vector>

namespace Preview
	{
	struct Line
		{
		enum Kind
			{
			Section,	//!< a heading across the list: text
			Op,			//!< one operation: text (op and type), detail (its comment)
			Change,		//!< one value: text (column), from, to
			Refused,	//!< one refused row: text (where), detail (why)
			Note		//!< plain information: text, detail
			};
		enum Box { NoBox, Ticked, Unticked, Mixed };

		Kind kind = Note;
		std::wstring text;
		std::wstring detail;
		std::wstring from;
		std::wstring to;

		/// Op and Change lines: the tick box. An Op's box follows its changes
		/// (Mixed when some are left out).
		Box box = NoBox;
		/// Change: changes of one operation with the same link tick together.
		std::wstring link;
		/// The caller's own number: which change, which operation.
		long tag = 0;
		/// Op: its effect on time and flips, shown at the right; `tone` -1
		/// better (green), +1 worse (red), 0 plain.
		std::wstring impact;
		int tone = 0;
		};

	/// Tick or untick line `i` and what goes with it - an Op line takes all its
	/// changes, a Change line the changes of its operation that share its link
	/// - then bring every Op box in line with its changes.
	void Toggle (std::vector<Line> &lines, size_t i);

	/// Every Op box set from its changes: all ticked, none, or Mixed.
	void Sync (std::vector<Line> &lines);

	/// How many Change lines are ticked.
	int CountTicked (const std::vector<Line> &lines);

	/// The Op line a Change line belongs to, or `lines.size ()`.
	size_t OpOf (const std::vector<Line> &lines, size_t i);

	struct Options
		{
		std::wstring caption = L"Parameter Table Tool - load preview";
		std::wstring verb = L"Apply";		//!< the button: "Apply 3 changes"
		std::wstring foot;					//!< under the list
		std::wstring nothing = L"Nothing would be written.";

		/// Run at the start and after every tick: sets each Op line's impact
		/// and returns the line for the top of the window (`tone` as on a
		/// line). Empty = no such line.
		std::function<std::wstring (std::vector<Line> &, int &tone)> impact;
		};

	/// Show the preview. With no ticked changes the window only offers Close.
	/// True when the person chose Apply with at least one change ticked; the
	/// lines' boxes then say which.
	bool Show (const std::wstring &title, const std::wstring &summary,
			   std::vector<Line> &lines, const Options &options);
	}
