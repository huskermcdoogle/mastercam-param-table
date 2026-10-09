#include "Plan.h"

#include <cwctype>
#include <map>
#include <set>

namespace Plan
	{
	int Schema::Find (const std::wstring &name) const
		{
		for (size_t i = 0; i < cols.size (); ++i)
			if (cols[i].name == name)
				return static_cast<int> (i);
		return -1;
		}

	bool IsInfoColumn (const std::wstring &name)
		{
		static const wchar_t *const info[] = {
			L"tool", L"comment", L"description", L"tool_radius", L"tool_name",
			L"group_name", L"units", L"needs_regen", L"coolant_text", L"coolant_x", L"canned_text_raw", L"changes",
			L"cycle_time", L"cycle_time_raw", L"travel_x_min", L"travel_x_max", L"travel_z_min",
			L"travel_z_max", L"cut_length", L"rapid_length", L"feed_groups",
			L"est_cycle_time", L"time_change", L"est_seconds",
			L"flips", L"flips_est", L"flips_why", L"flip_longest", L"insp_mode", L"cut_seconds_est",
			L"cut_dia", L"mrr", L"mrr_avg", L"removed_est", L"mrr_basis",
			L"flips_uncommented", L"flips_part", L"edge_limit", L"edge_after" };
		for (const wchar_t *n : info)
			if (name == n)
				return true;
		return false;
		}

	std::wstring Lower (std::wstring s)
		{
		for (wchar_t &c : s)
			c = static_cast<wchar_t> (std::towlower (c));
		return s;
		}

	bool SameValue (const Col &col, const std::wstring &a, const std::wstring &b)
		{
		switch (col.type)
			{
			case Type::Bool:
				{
				bool x = false, y = false;
				return Csv::ParseBool (a, x) && Csv::ParseBool (b, y) && x == y;
				}
			case Type::Long:
				{
				long long x = 0, y = 0;
				return Csv::ParseLong (a, x) && Csv::ParseLong (b, y) && x == y;
				}
			case Type::Double:
				{
				double x = 0, y = 0;
				return Csv::ParseDouble (a, x) && Csv::ParseDouble (b, y)
					   && Csv::SameDouble (x, y);
				}
			case Type::Text:
				// A choice is the same choice whatever its case.
				if (!col.choices.empty ())
					return Lower (Csv::Trim (a)) == Lower (Csv::Trim (b));
				return a == b;
			}
		return false;
		}

	bool Valid (const Col &col, const std::wstring &text, std::wstring &why)
		{
		switch (col.type)
			{
			case Type::Bool:
				{
				bool b = false;
				if (!Csv::ParseBool (text, b))
					{
					why = L"\"" + text + L"\" is not 1/0 or true/false";
					return false;
					}
				return true;
				}
			case Type::Long:
				{
				long long n = 0;
				if (!Csv::ParseLong (text, n))
					{
					why = L"\"" + text + L"\" is not a whole number";
					return false;
					}
				const double v = static_cast<double> (n);
				if (!std::isnan (col.lo) && v < col.lo)
					{
					why = L"\"" + text + L"\" is below the minimum "
						  + Csv::FormatDouble (col.lo);
					return false;
					}
				if (!std::isnan (col.hi) && v > col.hi)
					{
					why = L"\"" + text + L"\" is above the maximum "
						  + Csv::FormatDouble (col.hi);
					return false;
					}
				return true;
				}
			case Type::Double:
				{
				double v = 0;
				if (!Csv::ParseDouble (text, v))
					{
					why = L"\"" + text + L"\" is not a number";
					return false;
					}
				if (!std::isnan (col.lo) && v < col.lo)
					{
					why = L"\"" + text + L"\" is below the minimum "
						  + Csv::FormatDouble (col.lo);
					return false;
					}
				if (!std::isnan (col.hi) && v > col.hi)
					{
					why = L"\"" + text + L"\" is above the maximum "
						  + Csv::FormatDouble (col.hi);
					return false;
					}
				return true;
				}
			case Type::Text:
				// The operation's own buffer has a fixed size, and a longer
				// cell would be CUT OFF when written - a comment that quietly
				// loses its tail. Refuse it instead, and say by how much.
				if (!std::isnan (col.hi)
					&& static_cast<double> (text.size ()) > col.hi)
					{
					why = L"is " + std::to_wstring (text.size ())
						  + L" characters; the operation holds at most "
						  + Csv::FormatDouble (col.hi);
					return false;
					}
				if (!col.choices.empty ())
					{
					for (const std::wstring &c : col.choices)
						if (Lower (c) == Lower (Csv::Trim (text)))
							return true;
					why = L"\"" + text + L"\" is not one of:";
					for (const std::wstring &c : col.choices)
						why += L" \"" + c + L"\"";
					return false;
					}
				if (col.check != nullptr && !col.check (text, why))
					return false;
				return true;
			}
		return true;
		}

	Result Make (const Schema &schema, const std::vector<Current> &current,
				 const std::vector<Csv::Row> &csv)
		{
		Result r;

		auto refuse = [&r] (size_t line, long op, const std::wstring &why)
			{
			Refusal f;
			f.line = line;
			f.op = op;
			f.why = why;
			r.refusals.push_back (f);
			};

		if (csv.empty ())
			{
			refuse (1, 0, L"the file is empty");
			return r;
			}

		// ---- The header decides which cell means what.
		const Csv::Row &head = csv[0];
		int opCol = -1, typeCol = -1;
		std::map<std::wstring, size_t> colOf;		// header name -> position
		for (size_t i = 0; i < head.size (); ++i)
			{
			std::wstring h = head[i];
			const size_t a = h.find_first_not_of (L" \t");
			const size_t b = h.find_last_not_of (L" \t");
			h = (a == std::wstring::npos) ? L"" : h.substr (a, b - a + 1);

			if (h == L"op_idn")
				opCol = static_cast<int> (i);
			else if (h == L"type")
				typeCol = static_cast<int> (i);

			if (colOf.count (h) != 0 && !h.empty ())
				{
				refuse (1, 0, L"the column \"" + h + L"\" appears twice");
				return r;
				}
			colOf[h] = i;

			if (!h.empty () && h != L"op_idn" && h != L"type" && !IsInfoColumn (h)
				&& schema.Find (h) < 0)
				r.unknownColumns.push_back (h);
			}

		if (opCol < 0 || typeCol < 0)
			{
			refuse (1, 0, L"the header has no op_idn and type columns, so no row "
						  L"can be matched to an operation");
			return r;
			}

		std::map<long, const Current *> byOp;
		for (const Current &c : current)
			byOp[c.op] = &c;

		std::set<long> seen;

		// ---- The data rows.
		for (size_t line = 1; line < csv.size (); ++line)
			{
			const Csv::Row &row = csv[line];
			const size_t lineNo = line + 1;
			++r.rows;

			auto cell = [&] (int col) -> std::wstring
				{
				return (col >= 0 && static_cast<size_t> (col) < row.size ())
						   ? row[static_cast<size_t> (col)] : std::wstring ();
				};

			long long op = 0;
			if (!Csv::ParseLong (cell (opCol), op) || op <= 0)
				{
				refuse (lineNo, 0, L"op_idn \"" + cell (opCol) + L"\" is not an "
								   L"operation number");
				continue;
				}

			const auto found = byOp.find (static_cast<long> (op));
			if (found == byOp.end ())
				{
				refuse (lineNo, static_cast<long> (op),
						L"no " + schema.type + L" operation with that number is "
						L"in the part");
				continue;
				}
			const Current &cur = *found->second;

			// The type check is what stops a rough row landing on a finish op.
			if (cell (typeCol) != schema.type || cur.type != schema.type)
				{
				refuse (lineNo, static_cast<long> (op),
						L"the row says type \"" + cell (typeCol) + L"\" but "
						L"operation " + std::to_wstring (op) + L" is a "
						L"different kind of operation");
				continue;
				}

			if (!seen.insert (static_cast<long> (op)).second)
				{
				refuse (lineNo, static_cast<long> (op),
						L"operation " + std::to_wstring (op) + L" appears more "
						L"than once - which row is meant?");
				continue;
				}

			// ---- A value in a column that belongs to another kind of operation.
			// The dump leaves these blank, so anything here was typed - and it
			// would change nothing, which the person would not know.
			bool foreignHit = false;
			for (const std::wstring &f : schema.foreign)
				{
				const auto at = colOf.find (f);
				if (at == colOf.end ())
					continue;
				const std::wstring text = cell (static_cast<int> (at->second));
				if (text.find_first_not_of (L" \t") == std::wstring::npos)
					continue;
				refuse (lineNo, static_cast<long> (op),
						L"\"" + f + L"\" does not apply to a " + schema.type
						+ L" operation, so \"" + text + L"\" would change nothing "
						L"- clear the cell or use the right column");
				foreignHit = true;
				}
			if (foreignHit)
				continue;

			// ---- Judge every cell BEFORE recording any change from the row,
			// so a bad cell refuses the whole row rather than half of it.
			std::vector<Change> rowChanges;
			bool bad = false;

			for (size_t k = 0; k < schema.cols.size (); ++k)
				{
				const Col &col = schema.cols[k];
				const auto at = colOf.find (col.name);
				if (at == colOf.end ())
					continue;		// this sheet does not carry that column

				const std::wstring text = cell (static_cast<int> (at->second));
				const std::wstring have = k < cur.values.size () ? cur.values[k]
																  : std::wstring ();

				// Whitespace-only counts as empty: Excel pads nothing, but a
				// person's space bar is a real thing.
				if (text.find_first_not_of (L" \t") == std::wstring::npos
					&& col.type != Type::Text)
					{
					++r.emptyCells;
					continue;
					}
				if (text.empty ())
					{
					++r.emptyCells;
					continue;
					}

				if (SameValue (col, text, have))
					continue;

				if (col.readOnly)
					{
					refuse (lineNo, static_cast<long> (op),
							L"\"" + col.name + L"\" is read-only here but was "
							L"changed from " + have + L" to " + text);
					bad = true;
					continue;
					}

				std::wstring why;
				if (!Valid (col, text, why))
					{
					refuse (lineNo, static_cast<long> (op),
							L"\"" + col.name + L"\": " + why);
					bad = true;
					continue;
					}

				Change c;
				c.op = static_cast<long> (op);
				c.col = k;
				c.name = col.name;
				c.from = have;
				c.to = text;
				rowChanges.push_back (c);
				}

			if (bad)
				continue;

			if (rowChanges.empty ())
				++r.rowsUnchanged;
			else
				{
				++r.rowsChanged;
				r.changes.insert (r.changes.end (), rowChanges.begin (),
								  rowChanges.end ());
				}
			}

		return r;
		}
	}
