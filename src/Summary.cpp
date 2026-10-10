#include "Summary.h"
#include "Csv.h"

#include <algorithm>
#include <cmath>
#include <cwchar>
#include <functional>

namespace
	{
	using Cell = Xlsx::Sheet::FreeCell;

	/// The page's columns: A rank, B what (comment / name / label), C to H the
	/// figures, I the hidden working column of the ranked lists.
	enum { CA, CB, CC, CD, CE, CF, CG, CH, CI };
	const double kWidths[] = { 4, 46, 12, 16, 12, 13, 13, 12, 0 };

	const wchar_t *const kMain = L"'Lathe params'!";
	const wchar_t *const kTools = L"'Tools'!";
	const wchar_t *const kHmsFmt = L"\"[h]:mm:ss\"";

	std::wstring Letters (size_t col)
		{
		const std::string l = Xlsx::ColName (col);
		return std::wstring (l.begin (), l.end ());
		}

	/// Seconds as h:mm:ss (as Excel's [h]:mm:ss shows them).
	std::wstring Hms (double seconds)
		{
		if (!(seconds >= 0))
			return std::wstring ();
		const long long t = static_cast<long long> (seconds + 0.5);
		wchar_t buf[32];
		swprintf_s (buf, L"%lld:%02lld:%02lld", t / 3600, (t / 60) % 60, t % 60);
		return buf;
		}

	Cell Text (const std::wstring &t, Cell::Look look = Cell::Plain)
		{
		Cell c;
		c.text = t;
		c.look = look;
		return c;
		}

	/// An empty cell of the page: nothing drawn.
	Cell Blank ()
		{
		return Text (L"");
		}

	Cell Head (const std::wstring &t)
		{
		Cell c = Text (t);
		c.head = true;
		return c;
		}

	Cell Fx (const std::wstring &f, const std::wstring &cached, const std::wstring &fmt = L"",
			 Cell::Look look = Cell::Plain)
		{
		Cell c;
		c.formula = f;
		c.text = cached;
		c.numFmt = fmt;
		c.look = look;
		return c;
		}

	/// The same cell, aligned right: a figure under its heading.
	Cell Right (Cell c)
		{
		c.right = true;
		return c;
		}

	/// A number rounded as the sheet's ROUND(...,2) would show it.
	std::wstring Round2 (double v)
		{
		return Csv::Tidy (std::round (v * 100.0) / 100.0);
		}

	/// The page, row by row.
	struct Page
		{
		std::vector<std::vector<Cell>> rows;
		size_t Next () const { return rows.size () + 1; }		//!< the sheet row the next Add lands on
		void Add (std::vector<Cell> r = {}) { rows.push_back (std::move (r)); }
		};
	}

namespace Summary
	{
	std::vector<std::wstring> CostHeads ()
		{
		return { L"Cost per insert", L"Insert cost / part" };
		}

	std::vector<Xlsx::Sheet::FreeCell> CostCells (size_t row)
		{
		const std::wstring r = std::to_wstring (row);
		Cell cost;
		cost.editable = true;
		cost.numFmt = L"currency";
		Cell per;
		per.numFmt = L"currency";
		per.formula = L"IF(AND(N($D" + r + L")>0,ISNUMBER($E" + r + L"),ISNUMBER($G" + r + L")),$E" + r + L"/$D" + r
					  + L"*$G" + r + L",\"\")";
		return { cost, per };
		}

	void Add (Xlsx::Sheet &s, const Where &w)
		{
		if (s.rows.empty ())
			return;
		const std::vector<std::wstring> &names = s.rows[0];
		auto colOf = [&names] (const wchar_t *n) -> int
			{
			for (size_t c = 0; c < names.size (); ++c)
				if (names[c] == n)
					return static_cast<int> (c);
			return -1;
			};
		const int opCol = colOf (L"op_idn"), typeCol = colOf (L"type"), toolCol = colOf (L"tool");
		const int commentCol = colOf (L"comment"), estCol = colOf (L"est_seconds"), cutCol = colOf (L"cut_seconds_est");
		const int flipsCol = colOf (L"flips_part"), remCol = colOf (L"removed"), airCol = colOf (L"air_pct");
		const size_t nOps = s.rows.size () - 1;
		const std::wstring last = std::to_wstring (nOps + 2);

		// A main-sheet column as an absolute range, live or as dumped.
		auto range = [&] (int c, bool dumped = false)
			{
			const std::wstring l = Letters (static_cast<size_t> (c));
			return std::wstring (dumped ? L"Dumped!" : kMain) + L"$" + l + L"$3:$" + l + L"$" + last;
			};
		auto num = [&] (size_t d, int c, double &v)
			{
			return c >= 0 && d + 1 < s.rows.size () && static_cast<size_t> (c) < s.rows[d + 1].size ()
				   && Csv::ParseDouble (s.rows[d + 1][static_cast<size_t> (c)], v);
			};
		auto text = [&] (size_t d, int c)
			{
			return c >= 0 && d + 1 < s.rows.size () && static_cast<size_t> (c) < s.rows[d + 1].size ()
					   ? s.rows[d + 1][static_cast<size_t> (c)] : std::wstring ();
			};
		auto sumOf = [&] (int c)
			{
			double t = 0, v = 0;
			for (size_t d = 0; d < nOps; ++d)
				if (num (d, c, v))
					t += v;
			return t;
			};

		Page p;
		p.Add ({ Text (w.title, Cell::Title) });
		if (!w.subtitle.empty ())
			p.Add ({ Text (w.subtitle, Cell::Note) });
		p.Add ();

		// ---- How to use it: four lines, the whole manual.
		p.Add ({ Text (L"How to use this workbook", Cell::Section) });
		for (const wchar_t *line : {
				 L"1. Edit feeds, speeds and depths on 'Lathe params'. Edited cells turn yellow; worked-out figures that move turn blue.",
				 L"2. This page follows every edit: the part's totals as dumped and now, and where the time and inserts go.",
				 L"3. On 'Tools', type the edges and cost per insert; type the batch quantity below for the cost of a batch.",
				 L"4. Save, then run \"Lathe params - load\" in Mastercam: it shows every change before writing any." })
			p.Add ({ Text (line, Cell::Note) });
		// The user manual: a link that works with or without macros.
		if (!w.manual.empty ())
			{
			std::wstring path;
			for (wchar_t ch : w.manual)
				path += ch == L'"' ? std::wstring (L"\"\"") : std::wstring (1, ch);
			p.Add ({ Fx (L"HYPERLINK(\"" + path + L"\",\"Open the user manual - how every tool works, with examples\")",
						 L"Open the user manual - how every tool works, with examples", L"", Cell::Link) });
			}
		p.Add ();

		// ---- PART TOTALS, as dumped and now: the effect of the edits on the whole
		// part. "As dumped" sums the hidden Dumped sheet; "Now" the live sheet.
		if (estCol >= 0)
			{
			p.Add ({ Text (L"Part totals", Cell::Section) });
			p.Add ({ Blank (), Head (L"What"), Right (Head (L"As dumped")), Right (Head (L"Now")),
					 Right (Head (L"Change")) });
			auto sum = [&] (int c, bool dumped) { return L"SUM(" + range (c, dumped) + L")"; };
			auto timeRow = [&] (const wchar_t *what, int c)
				{
				const std::wstring was = sum (c, true), now = sum (c, false), v0 = Hms (sumOf (c));
				p.Add ({ Blank (), Text (what),
						 Right (Fx (L"TEXT(" + was + L"/86400," + kHmsFmt + L")", v0)),
						 Right (Fx (L"TEXT(" + now + L"/86400," + kHmsFmt + L")", v0)),
						 Right (Fx (L"IF(ROUND(" + now + L"-" + was + L",0)<0,\"-\",\"+\")&TEXT(ABS(ROUND(" + now + L"-" + was
									+ L",0))/86400," + kHmsFmt + L")", L"+0:00:00")) });
				};
			auto numRow = [&] (const wchar_t *what, const std::wstring &was, const std::wstring &now, double v0)
				{
				p.Add ({ Blank (), Text (what), Right (Fx (L"ROUND(" + was + L",2)", Round2 (v0))),
						 Right (Fx (L"ROUND(" + now + L",2)", Round2 (v0))),
						 Right (Fx (L"ROUND(" + now + L"-(" + was + L"),2)", L"0")) });
				};
			timeRow (L"Cycle time (estimate)", estCol);
			if (cutCol >= 0)
				timeRow (L"Cutting time", cutCol);
			if (flipsCol >= 0)
				numRow (L"Insert flips per part", sum (flipsCol, true), sum (flipsCol, false), sumOf (flipsCol));
			if (remCol >= 0)
				{
				numRow (L"Material removed (fixed by the toolpaths)", sum (remCol, true), sum (remCol, false), sumOf (remCol));
				const double est0 = sumOf (estCol);
				numRow (L"Average removal rate over the cycle (per min)",
						L"IFERROR(" + sum (remCol, true) + L"/" + sum (estCol, true) + L"*60,0)",
						L"IFERROR(" + sum (remCol, false) + L"/" + sum (estCol, false) + L"*60,0)",
						est0 > 0 ? sumOf (remCol) / est0 * 60.0 : 0);
				}
			p.Add ();
			}

		// ---- A BATCH: the quantity typed in; time and insert cost for that many.
		{
		p.Add ({ Text (L"Batch", Cell::Section) });
		const size_t qtyRow = p.Next ();
		Cell qty = Text (w.batchQty.empty () ? std::wstring (L"1") : w.batchQty, Cell::Input);
		qty.editable = true;
		p.Add ({ Blank (), Text (L"Batch quantity (type it in)"), qty });
		{
		Xlsx::Sheet::Validation v;
		v.cells = "C" + std::to_string (qtyRow);
		v.type = "whole";
		v.op = "greaterThanOrEqual";
		v.f1 = L"1";
		v.title = L"Batch quantity";
		v.prompt = L"How many parts in a batch - for the cycle time, flips and whole inserts per batch.";
		v.error = L"A whole number of parts, 1 or more.";
		s.summaryValidations.push_back (v);
		}
		const std::wstring q = L"$C$" + std::to_wstring (qtyRow);
		if (estCol >= 0)
			p.Add ({ Blank (), Text (L"Cycle time per batch (now)"),
					 Right (Fx (L"TEXT(SUM(" + range (estCol) + L")*" + q + L"/86400," + kHmsFmt + L")", Hms (sumOf (estCol)))) });
		if (flipsCol >= 0)
			p.Add ({ Blank (), Text (L"Insert flips per batch (now)"),
					 Right (Fx (L"ROUND(SUM(" + range (flipsCol) + L")*" + q + L",2)", Round2 (sumOf (flipsCol)))) });
		// The inserts table: its heading row, then rows to the first empty one.
		size_t insertFirst = 0, insertLast = 0;
		const size_t after0 = s.tools.size () + 3;		// the Tools page row of toolsAfter[0]
		for (size_t k = 0; k < s.toolsAfter.size () && insertFirst == 0; ++k)
			{
			const std::vector<Cell> &row = s.toolsAfter[k];
			if (row.size () > CG && row[CA].head && row[CA].text == L"Inserts" && row[CG].text == CostHeads ()[0])
				{
				insertFirst = after0 + k + 1;
				for (size_t e = k + 1; e < s.toolsAfter.size () && !s.toolsAfter[e].empty (); ++e)
					insertLast = after0 + e;
				}
			}
		if (insertFirst > 0 && insertLast >= insertFirst)
			{
			auto tcol = [&] (const wchar_t *l)
				{
				return std::wstring (kTools) + L"$" + l + L"$" + std::to_wstring (insertFirst) + L":$" + l + L"$"
					   + std::to_wstring (insertLast);
				};
			// What was typed at the dump (nothing, usually), for the cached figures.
			double perPart0 = 0, batch0 = 0;
			for (size_t r = insertFirst; r <= insertLast; ++r)
				{
				if (r < after0 || r - after0 >= s.toolsAfter.size ())
					continue;
				const std::vector<Cell> &row = s.toolsAfter[r - after0];
				double edges = 0, flips = 0, cost = 0;
				if (row.size () > CG && Csv::ParseDouble (row[CD].text, edges) && edges > 0
					&& Csv::ParseDouble (row[CE].text, flips) && Csv::ParseDouble (row[CG].text, cost))
					{
					perPart0 += flips / edges * cost;
					batch0 += std::ceil (std::round (flips / edges * 1e6) / 1e6) * cost;
					}
				}
			p.Add ({ Blank (), Text (L"Insert cost per part (average)"),
					 Fx (L"SUM(" + tcol (L"H") + L")", Round2 (perPart0), L"currency") });
			// Whole inserts for the batch: per insert, ROUNDUP(flips x qty / edges).
			Cell batch = Fx (L"SUM(IFERROR(ROUNDUP(ROUND(" + tcol (L"E") + L"*" + q + L"/" + tcol (L"D") + L",6),0)*"
							 + tcol (L"G") + L",0))", Round2 (batch0), L"currency");
			batch.array = true;
			p.Add ({ Blank (), Text (L"Insert cost per batch (whole inserts)"), batch });

			// ---- INSERT USAGE, per insert: what a part (and the batch) uses of
			// each, straight off the Tools page's inserts table - live, including
			// the rows typed into there. Blank names stay blank.
			p.Add ();
			p.Add ({ Text (L"Insert usage per part", Cell::Section) });
			p.Add ({ Blank (), Head (L"Insert"), Head (L"Used by"), Right (Head (L"Edges")), Right (Head (L"Flips / part")),
					 Right (Head (L"Inserts / part")), Right (Head (L"Inserts / batch")), Right (Head (L"Cost / part")) });
			for (size_t r = insertFirst; r <= insertLast; ++r)
				{
				if (r < after0 || r - after0 >= s.toolsAfter.size ())
					continue;
				const std::vector<Cell> &row = s.toolsAfter[r - after0];
				auto cellOf = [&] (size_t c) { return c < row.size () ? row[c].text : std::wstring (); };
				const std::wstring R = std::to_wstring (r);
				auto ref = [&] (const wchar_t *l) { return std::wstring (kTools) + L"$" + l + L"$" + R; };
				auto shown = [&] (const wchar_t *l, size_t c, bool number)
					{
					// Blank while the insert row is unnamed; else the Tools cell.
					Cell x = Fx (L"IF(" + ref (L"B") + L"=\"\",\"\"," + ref (l) + L")", cellOf (c));
					return number ? Right (x) : x;
					};
				double edges = 0, flips = 0;
				const bool known = Csv::ParseDouble (cellOf (CD), edges) && edges > 0 && Csv::ParseDouble (cellOf (CE), flips);
				Cell batchN = Right (Fx (L"IF(OR(" + ref (L"B") + L"=\"\",N(" + ref (L"D") + L")<=0),\"\",ROUNDUP(ROUND(N("
										 + ref (L"E") + L")*" + q + L"/" + ref (L"D") + L",6),0))",
										 known ? Csv::Tidy (std::ceil (std::round (flips / edges * 1e6) / 1e6)) : std::wstring ()));
				p.Add ({ Blank (), shown (L"B", CB, false), shown (L"C", CC, false), shown (L"D", CD, true),
						 shown (L"E", CE, true), shown (L"F", CF, true), batchN,
						 Fx (L"IF(" + ref (L"B") + L"=\"\",\"\"," + ref (L"H") + L")", cellOf (CH), L"currency") });
				}
			}
		p.Add ();
		}

		// ---- THE RANKED LISTS. `key` is the array formula of the value per main
		// sheet row; key0 the dumped value per row (NaN = none). A row shows when
		// its value is above 0.
		auto ranked = [&] (size_t count, const std::wstring &key, const std::vector<double> &key0,
						   const std::function<std::vector<Cell> (size_t k, const std::wstring &at, long pos0)> &cells)
			{
			std::vector<size_t> order;
			for (size_t d = 0; d < key0.size (); ++d)
				if (key0[d] > 0)
					order.push_back (d);
			std::stable_sort (order.begin (), order.end (), [&key0] (size_t a, size_t b) { return key0[a] > key0[b]; });
			for (size_t k = 1; k <= count; ++k)
				{
				const std::wstring at = L"$I" + std::to_wstring (p.Next ());
				const long pos0 = k <= order.size () ? static_cast<long> (order[k - 1]) + 1 : -1;
				std::vector<Cell> row = cells (k, at, pos0);
				row.insert (row.begin (), Fx (L"IF(" + at + L"=\"\",\"\"," + std::to_wstring (k) + L")",
											  pos0 > 0 ? std::to_wstring (k) : std::wstring (), L"", Cell::Note));
				row.resize (CI, Blank ());
				const std::wstring n = std::to_wstring (k), large = L"LARGE(" + key + L"," + n + L")";
				Cell helper = Fx (L"IFERROR(IF(" + large + L"<0,\"\",MATCH(" + large + L"," + key + L",0)),\"\")",
								  pos0 > 0 ? std::to_wstring (pos0) : std::wstring ());
				helper.array = true;
				row.push_back (helper);
				p.Add (row);
				}
			p.Add ();
			};
		// A main-sheet cell of the n-th row: text, op number (a link to the row), time.
		auto mainText = [&] (int c, const std::wstring &at, long pos0, const wchar_t *none)
			{
			if (c < 0)
				return Blank ();
			const std::wstring f = L"INDEX(" + range (c) + L"," + at + L")&\"\"";
			return Fx (L"IF(" + at + L"=\"\"," + (none ? L"\"" + std::wstring (none) + L"\"" : L"\"\"") + L"," + f + L")",
					   pos0 > 0 ? text (static_cast<size_t> (pos0 - 1), c) : none ? none : L"");
			};
		auto opLink = [&] (const std::wstring &at, long pos0)
			{
			if (opCol < 0)
				return Blank ();
			return Fx (L"IF(" + at + L"=\"\",\"\",HYPERLINK(\"#'Lathe params'!A\"&(" + at + L"+2),INDEX(" + range (opCol) + L","
					   + at + L")&\"\"))", pos0 > 0 ? text (static_cast<size_t> (pos0 - 1), opCol) : L"", L"", Cell::Link);
			};
		auto rowKey = [&] (int c)
			{
			std::vector<double> k (nOps, std::nan (""));
			for (size_t d = 0; d < nOps; ++d)
				num (d, c, k[d]);
			return k;
			};
		const std::wstring tieBreak = L"-ROW(" + range (opCol >= 0 ? opCol : 0) + L")/1E9";
		const size_t topOps = (std::min) (kTop, nOps);

		// Longest operations.
		if (estCol >= 0 && topOps > 0)
			{
			p.Add ({ Text (L"Longest operations (estimated time, now)", Cell::Section) });
			p.Add ({ Head (L"#"), Head (L"Comment"), Head (L"Op"), Head (L"Type"), Head (L"Tool"), Right (Head (L"Time")),
					 Right (Head (L"Share of part")) });
			const std::wstring est = range (estCol);
			const double total0 = sumOf (estCol);
			const std::vector<double> key0 = rowKey (estCol);
			ranked (topOps, L"IF(ISNUMBER(" + est + L")," + est + L",-1)" + tieBreak, key0,
					[&] (size_t k, const std::wstring &at, long pos0)
				{
				const double v0 = pos0 > 0 ? key0[static_cast<size_t> (pos0 - 1)] : 0;
				const std::wstring idx = L"INDEX(" + est + L"," + at + L")";
				return std::vector<Cell> {
					mainText (commentCol, at, pos0, k == 1 ? L"No estimated times - regenerate and dump again." : nullptr),
					opLink (at, pos0), mainText (typeCol, at, pos0, nullptr), mainText (toolCol, at, pos0, nullptr),
					Right (Fx (L"IF(" + at + L"=\"\",\"\",TEXT(" + idx + L"/86400," + kHmsFmt + L"))", pos0 > 0 ? Hms (v0) : L"")),
					Right (Fx (L"IF(" + at + L"=\"\",\"\",IFERROR(" + idx + L"/SUM(" + est + L"),\"\"))",
							   pos0 > 0 && total0 > 0 ? Csv::Tidy (v0 / total0) : L"", L"0.0%")) };
				});
			}

		// Tools by insert flips: from the Tools page's own live sums.
		const auto &heads = s.toolExtraHeads;
		auto extraCol = [&heads] (const wchar_t *h) -> int
			{
			for (size_t e = 0; e < heads.size (); ++e)
				if (heads[e] == h)
					return static_cast<int> (3 + e);
			return -1;
			};
		const int tFlips = extraCol (L"Flips / part"), tCut = extraCol (L"Cut time / part"), tInsert = extraCol (L"Insert");
		if (!s.tools.empty () && tFlips >= 0)
			{
			const std::wstring lastTool = std::to_wstring (s.tools.size () + 1);
			auto trange = [&] (int c)
				{
				const std::wstring l = Letters (static_cast<size_t> (c));
				return std::wstring (kTools) + L"$" + l + L"$2:$" + l + L"$" + lastTool;
				};
			auto extra = [&] (long pos0, int c)
				{
				if (pos0 <= 0 || c < 3)
					return std::wstring ();
				const auto &x = s.tools[static_cast<size_t> (pos0 - 1)].extra;
				return static_cast<size_t> (c - 3) < x.size () ? x[static_cast<size_t> (c - 3)].text : std::wstring ();
				};
			std::vector<double> key0 (s.tools.size (), std::nan (""));
			for (size_t k = 0; k < s.tools.size (); ++k)
				Csv::ParseDouble (extra (static_cast<long> (k) + 1, tFlips), key0[k]);
			p.Add ({ Text (L"Insert flips by tool (per part, now)", Cell::Section) });
			p.Add ({ Head (L"#"), Head (L"Tool name"), Head (L"Tool"), Head (L"Insert"), Right (Head (L"Flips / part")),
					 Right (Head (L"Cut time / part")) });
			const std::wstring flips = trange (tFlips);
			auto toolText = [&] (int c, const std::wstring &at, long pos0, const wchar_t *none)
				{
				const std::wstring f = L"INDEX(" + trange (c) + L"," + at + L")&\"\"";
				std::wstring v0 = none ? none : L"";
				if (pos0 > 0)
					v0 = c == 1 ? s.tools[static_cast<size_t> (pos0 - 1)].name : extra (pos0, c);
				return Fx (L"IF(" + at + L"=\"\"," + (none ? L"\"" + std::wstring (none) + L"\"" : L"\"\"") + L"," + f + L")", v0);
				};
			ranked ((std::min) (kTop, s.tools.size ()), L"IF(ISNUMBER(" + flips + L")," + flips + L",-1)-ROW(" + flips + L")/1E9",
					key0, [&] (size_t k, const std::wstring &at, long pos0)
				{
				return std::vector<Cell> {
					toolText (1, at, pos0, k == 1 ? L"No insert flips counted on this part." : nullptr),
					Fx (L"IF(" + at + L"=\"\",\"\",HYPERLINK(\"#'Tools'!A\"&(" + at + L"+1),INDEX(" + trange (0) + L"," + at + L")&\"\"))",
						pos0 > 0 ? s.tools[static_cast<size_t> (pos0 - 1)].number : L"", L"", Cell::Link),
					tInsert >= 0 ? toolText (tInsert, at, pos0, nullptr) : Blank (),
					Right (Fx (L"IF(" + at + L"=\"\",\"\",ROUND(INDEX(" + flips + L"," + at + L"),2))",
							   pos0 > 0 ? Round2 (key0[static_cast<size_t> (pos0 - 1)]) : L"")),
					tCut >= 0 ? Right (toolText (tCut, at, pos0, nullptr)) : Blank () };
				});
			}

		// Air cutting: air_pct x cut time - the time spent cutting nothing.
		if (airCol >= 0 && cutCol >= 0 && topOps > 0)
			{
			const std::wstring air = range (airCol), cut = range (cutCol);
			std::vector<double> key0 (nOps, std::nan (""));
			for (size_t d = 0; d < nOps; ++d)
				{
				double a = 0, c = 0;
				if (num (d, airCol, a) && num (d, cutCol, c))
					key0[d] = a / 100.0 * c;
				}
			p.Add ({ Text (L"Most air cutting (cut time with no stock in the way)", Cell::Section) });
			p.Add ({ Head (L"#"), Head (L"Comment"), Head (L"Op"), Head (L"Type"), Head (L"Tool"), Right (Head (L"Air time")),
					 Right (Head (L"Air share")), Right (Head (L"Cut time")) });
			ranked (topOps, L"IF(ISNUMBER(" + air + L")*ISNUMBER(" + cut + L")," + air + L"/100*" + cut + L",-1)" + tieBreak, key0,
					[&] (size_t k, const std::wstring &at, long pos0)
				{
				double a0 = 0, c0 = 0;
				if (pos0 > 0)
					{
					num (static_cast<size_t> (pos0 - 1), airCol, a0);
					num (static_cast<size_t> (pos0 - 1), cutCol, c0);
					}
				const std::wstring ia = L"INDEX(" + air + L"," + at + L")", ic = L"INDEX(" + cut + L"," + at + L")";
				return std::vector<Cell> {
					mainText (commentCol, at, pos0, k == 1 ? L"No air cutting measured." : nullptr),
					opLink (at, pos0), mainText (typeCol, at, pos0, nullptr), mainText (toolCol, at, pos0, nullptr),
					Right (Fx (L"IF(" + at + L"=\"\",\"\",TEXT(" + ia + L"/100*" + ic + L"/86400," + kHmsFmt + L"))",
							   pos0 > 0 ? Hms (a0 / 100.0 * c0) : L"")),
					Right (Fx (L"IF(" + at + L"=\"\",\"\"," + ia + L"/100)", pos0 > 0 ? Csv::Tidy (a0 / 100.0) : L"", L"0%")),
					Right (Fx (L"IF(" + at + L"=\"\",\"\",TEXT(" + ic + L"/86400," + kHmsFmt + L"))", pos0 > 0 ? Hms (c0) : L"")) };
				});
			}

		s.summary = std::move (p.rows);
		s.summaryWidths.assign (std::begin (kWidths), std::end (kWidths));
		}
	}
