#include "Undo.h"

#include <cwctype>
#include <map>

namespace
	{
	const std::wstring kBegin = L"load begin: ";
	const std::wstring kFrom = L"  <-  ";
	const std::wstring kArrow = L" -> ";
	const std::wstring kBreak = L" ↵ ";

	std::wstring Lower (std::wstring s)
		{
		for (wchar_t &c : s)
			c = static_cast<wchar_t> (std::towlower (c));
		return s;
		}

	bool StartsWith (const std::wstring &s, const std::wstring &head)
		{
		return s.compare (0, head.size (), head) == 0;
		}

	/// A log line without its time stamp; `stamp` gets the stamp. Util::Log
	/// writes "YYYY-MM-DD HH:MM:SS" and two spaces before every line.
	std::wstring Unstamp (const std::wstring &line, std::wstring &stamp)
		{
		if (line.size () >= 21 && line[4] == L'-' && line[7] == L'-' && line[10] == L' '
			&& line[13] == L':' && line[16] == L':' && line.compare (19, 2, L"  ") == 0)
			{
			stamp = line.substr (0, 19);
			return line.substr (21);
			}
		stamp.clear ();
		return line;
		}

	/// "op 12 ROUGH  feed  0.3 -> 0.25" to an entry. False for any other line
	/// that starts with "op" - "op 12 FAILED - ..." among them.
	bool ParseEntry (const std::wstring &text, Undo::Entry &e)
		{
		if (!StartsWith (text, L"op "))
			return false;
		size_t i = 3;
		long op = 0;
		while (i < text.size () && iswdigit (text[i]))
			op = op * 10 + static_cast<long> (text[i++] - L'0');
		if (i == 3 || op <= 0 || i >= text.size () || text[i] != L' ')
			return false;
		++i;
		// The type may hold one space ("MILL DRILL"); two end it.
		const size_t t = text.find (L"  ", i);
		if (t == std::wstring::npos || t == i)
			return false;
		const size_t c = t + 2;
		const size_t ce = text.find (L"  ", c);
		if (ce == std::wstring::npos || ce == c)
			return false;
		const std::wstring column = text.substr (c, ce - c);
		if (column.find (L' ') != std::wstring::npos)
			return false;
		const std::wstring both = text.substr (ce + 2);
		if (both.find (kArrow) == std::wstring::npos && !StartsWith (both, kArrow.substr (1)))
			return false;
		e.op = op;
		e.type = text.substr (i, t - i);
		e.column = column;
		e.both = both;
		return true;
		}

	/// Whether a logged value is the value an operation holds now. The log has
	/// the text as typed into the sheet ("0.25"), the operation has it as
	/// stored ("0.25000000001", "Flood" for "flood") - equal as values counts.
	bool Same (const Plan::Col &col, const std::wstring &logged, const std::wstring &now)
		{
		const std::wstring shown = Undo::OneLine (now);
		if (shown == logged)
			return true;
		if (col.type == Plan::Type::Text)
			return Lower (Csv::Trim (shown)) == Lower (Csv::Trim (logged));
		return Plan::SameValue (col, logged, now);
		}
	}

namespace Undo
	{
	std::wstring BeginLine (const std::wstring &partName, const std::wstring &files)
		{
		return kBegin + partName + kFrom + files;
		}

	std::wstring ChangeLine (long op, const std::wstring &type, const std::wstring &column,
							 const std::wstring &from, const std::wstring &to)
		{
		return L"op " + std::to_wstring (op) + L" " + type + L"  " + column + L"  "
			   + OneLine (from) + kArrow + OneLine (to);
		}

	std::wstring OneLine (const std::wstring &v)
		{
		std::wstring o;
		for (wchar_t c : v)
			{
			if (c == L'\r')
				continue;
			o += c == L'\n' ? kBreak : std::wstring (1, c);
			}
		return o;
		}

	std::wstring FromOneLine (const std::wstring &v, const std::wstring &like)
		{
		const bool bareLf = like.find (L'\n') != std::wstring::npos
							&& like.find (L"\r\n") == std::wstring::npos;
		const std::wstring brk = bareLf ? L"\n" : L"\r\n";
		std::wstring o;
		size_t from = 0;
		for (;;)
			{
			const size_t at = v.find (kBreak, from);
			if (at == std::wstring::npos)
				break;
			o += v.substr (from, at - from) + brk;
			from = at + kBreak.size ();
			}
		return o + v.substr (from);
		}

	Last FindLast (const std::wstring &log, const std::wstring &partName)
		{
		Last best;
		bool ours = false;			// inside a load of this part
		bool anyLoad = false;		// inside a load of any part
		std::map<std::pair<long, std::wstring>, size_t> at;	// (op, column) -> entry
		const std::wstring want = Lower (partName);

		size_t from = 0;
		while (from < log.size ())
			{
			size_t eol = log.find (L'\n', from);
			if (eol == std::wstring::npos)
				eol = log.size ();
			std::wstring line = log.substr (from, eol - from);
			from = eol + 1;
			if (!line.empty () && line.back () == L'\r')
				line.pop_back ();
			if (!line.empty () && line[0] == L'﻿')
				line.erase (0, 1);

			std::wstring stamp;
			const std::wstring text = Unstamp (line, stamp);

			if (StartsWith (text, kBegin))
				{
				const std::wstring rest = text.substr (kBegin.size ());
				const size_t f = rest.find (kFrom);
				const std::wstring part = f == std::wstring::npos ? rest : rest.substr (0, f);
				anyLoad = true;
				ours = Lower (part) == want;
				if (ours)
					{
					// A newer load of this part: it is the one to undo.
					const bool older = best.olderLoads;
					best = Last ();
					best.found = true;
					best.olderLoads = older;
					best.stamp = stamp;
					best.files = f == std::wstring::npos ? std::wstring () : rest.substr (f + kFrom.size ());
					at.clear ();
					}
				continue;
				}
			if (StartsWith (text, L"load: "))
				{
				if (ours)
					best.end = text;
				else if (!anyLoad)
					best.olderLoads = true;
				ours = anyLoad = false;
				continue;
				}
			if (StartsWith (text, L"undo begin: "))
				{
				ours = anyLoad = false;
				continue;
				}

			Entry e;
			if (ours && ParseEntry (text, e))
				{
				// The same value twice in one load (two sheets naming one
				// operation): both were planned against the operation as it
				// stood before the load, and the later write is what is there.
				const auto key = std::make_pair (e.op, e.column);
				const auto had = at.find (key);
				if (had != at.end ())
					best.entries[had->second] = e;
				else
					{
					at[key] = best.entries.size ();
					best.entries.push_back (e);
					}
				}
			}
		return best;
		}

	std::vector<std::pair<std::wstring, std::wstring>> Splits (const std::wstring &both)
		{
		std::vector<std::pair<std::wstring, std::wstring>> out;
		// An empty old value leaves the line as " -> new" - its arrow at the start.
		if (StartsWith (both, kArrow.substr (1)))
			out.push_back ({ std::wstring (), both.substr (kArrow.size () - 1) });
		size_t from = 0;
		for (;;)
			{
			const size_t a = both.find (kArrow, from);
			if (a == std::wstring::npos)
				break;
			out.push_back ({ both.substr (0, a), both.substr (a + kArrow.size ()) });
			from = a + 1;
			}
		return out;
		}

	Result Make (const Last &last, const std::vector<Kind> &kinds)
		{
		Result r;

		// By operation, in the order the load wrote them.
		std::vector<long> order;
		std::map<long, std::vector<const Entry *>> byOp;
		for (const Entry &e : last.entries)
			{
			if (byOp.count (e.op) == 0)
				order.push_back (e.op);
			byOp[e.op].push_back (&e);
			}

		for (long op : order)
			{
			const std::vector<const Entry *> &es = byOp[op];
			const std::wstring &type = es.front ()->type;

			const Kind *kind = nullptr;
			for (const Kind &k : kinds)
				if (k.schema != nullptr && k.schema->type == type)
					kind = &k;
			const Plan::Current *cur = nullptr;
			if (kind != nullptr)
				for (const Plan::Current &c : kind->now)
					if (c.op == op)
						cur = &c;
			if (cur == nullptr)
				{
				Skip s;
				s.op = op;
				s.type = type;
				s.why = L"operation " + std::to_wstring (op) + L" is no longer a " + type
						+ L" operation in the part";
				r.skipped.push_back (s);
				continue;
				}

			// Every value of the operation judged before any is taken: one changed
			// since the load holds the whole operation back.
			std::vector<Restore> mine;
			std::vector<Skip> changed, unknown;
			for (const Entry *e : es)
				{
				const int col = kind->schema->Find (e->column);
				if (col < 0)
					{
					Skip s;
					s.op = op;
					s.type = type;
					s.column = e->column;
					s.why = L"\"" + e->column + L"\" is not a column this version writes";
					unknown.push_back (s);
					continue;
					}
				const Plan::Col &c = kind->schema->cols[static_cast<size_t> (col)];
				const std::wstring now = static_cast<size_t> (col) < cur->values.size ()
											 ? cur->values[static_cast<size_t> (col)] : std::wstring ();
				const auto splits = Splits (e->both);

				bool done = false;
				for (const auto &s : splits)
					if (Same (c, s.second, now))
						{
						Restore x;
						x.type = type;
						x.change.op = op;
						x.change.col = static_cast<size_t> (col);
						x.change.name = e->column;
						x.change.from = now;
						x.change.to = FromOneLine (s.first, now);
						mine.push_back (x);
						done = true;
						break;
						}
				if (done)
					continue;
				for (const auto &s : splits)
					if (Same (c, s.first, now))
						{
						r.already.push_back (*e);
						done = true;
						break;
						}
				if (done)
					continue;

				Skip s;
				s.op = op;
				s.type = type;
				s.column = e->column;
				s.why = L"\"" + e->column + L"\" changed since the load: it is "
						+ (now.empty () ? std::wstring (L"(empty)") : OneLine (now)) + L", the load wrote "
						+ (splits.empty () || splits.front ().second.empty () ? std::wstring (L"(empty)")
																				: splits.front ().second);
				changed.push_back (s);
				}

			r.skipped.insert (r.skipped.end (), unknown.begin (), unknown.end ());
			if (!changed.empty ())
				{
				r.skipped.insert (r.skipped.end (), changed.begin (), changed.end ());
				for (const Restore &x : mine)
					{
					Skip s;
					s.op = op;
					s.type = type;
					s.column = x.change.name;
					s.why = L"\"" + x.change.name + L"\" left as it is - another value of operation "
							+ std::to_wstring (op) + L" changed since the load";
					r.skipped.push_back (s);
					}
				continue;
				}
			r.restores.insert (r.restores.end (), mine.begin (), mine.end ());
			}
		return r;
		}
	}
