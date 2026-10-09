// PreviewTicks.cpp - the preview's tick boxes: what one click ticks. No
// Mastercam SDK and no window, so the rules are tested at a command prompt.
#include "Preview.h"

namespace Preview
	{
	size_t OpOf (const std::vector<Line> &lines, size_t i)
		{
		for (size_t k = i + 1; k-- > 0;)
			{
			if (k >= lines.size ())
				continue;
			if (lines[k].kind == Line::Op)
				return k;
			if (lines[k].kind != Line::Change)
				break;
			}
		return lines.size ();
		}

	void Sync (std::vector<Line> &lines)
		{
		for (size_t i = 0; i < lines.size (); ++i)
			{
			if (lines[i].kind != Line::Op || lines[i].box == Line::NoBox)
				continue;
			int on = 0, off = 0;
			for (size_t k = i + 1; k < lines.size () && lines[k].kind == Line::Change; ++k)
				{
				if (lines[k].box == Line::Ticked)
					++on;
				else if (lines[k].box == Line::Unticked)
					++off;
				}
			lines[i].box = off == 0 ? Line::Ticked : on == 0 ? Line::Unticked : Line::Mixed;
			}
		}

	void Toggle (std::vector<Line> &lines, size_t i)
		{
		if (i >= lines.size () || lines[i].box == Line::NoBox)
			return;

		if (lines[i].kind == Line::Op)
			{
			// A Mixed operation is ticked through, as a header box does in any
			// list: the click says "all of it", and a second click "none".
			const Line::Box to = lines[i].box == Line::Ticked ? Line::Unticked : Line::Ticked;
			for (size_t k = i + 1; k < lines.size () && lines[k].kind == Line::Change; ++k)
				if (lines[k].box != Line::NoBox)
					lines[k].box = to;
			}
		else if (lines[i].kind == Line::Change)
			{
			const Line::Box to = lines[i].box == Line::Ticked ? Line::Unticked : Line::Ticked;
			lines[i].box = to;
			if (!lines[i].link.empty ())
				{
				// Its partners sit under the same operation line.
				const size_t op = OpOf (lines, i);
				const size_t first = op < lines.size () ? op + 1 : i;
				for (size_t k = first; k < lines.size () && lines[k].kind == Line::Change; ++k)
					if (lines[k].link == lines[i].link && lines[k].box != Line::NoBox)
						lines[k].box = to;
				}
			}
		Sync (lines);
		}

	int CountTicked (const std::vector<Line> &lines)
		{
		int n = 0;
		for (const Line &l : lines)
			if (l.kind == Line::Change && l.box == Line::Ticked)
				++n;
		return n;
		}
	}
