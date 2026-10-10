#include "History.h"
#include "Summary.h"

#include <algorithm>
#include <ctime>
#include <cwchar>
#include <cwctype>
#include <fstream>
#include <iterator>
#include <system_error>

namespace
	{
	const std::wstring kSep = L" | ";
	const std::wstring kArrow = L" -> ";
	const double kNaN = std::nan ("");

	std::wstring Trim (const std::wstring &s)
		{
		const size_t a = s.find_first_not_of (L" \t\r\n");
		if (a == std::wstring::npos)
			return std::wstring ();
		return s.substr (a, s.find_last_not_of (L" \t\r\n") - a + 1);
		}

	std::wstring Lower (std::wstring s)
		{
		for (wchar_t &c : s)
			c = static_cast<wchar_t> (std::towlower (c));
		return s;
		}

	/// A typed text made fit for one field of one line: a line break shown the
	/// way the log shows one, a "|" (it splits the fields) as a broken bar.
	std::wstring OneField (const std::wstring &v)
		{
		std::wstring o;
		for (wchar_t c : v)
			{
			if (c == L'\r')
				continue;
			if (c == L'\n')
				o += L" ↵ ";
			else if (c == L'|')
				o += L'¦';
			else if (c == L'\t')
				o += L' ';
			else
				o += c;
			}
		return Trim (o);
		}

	/// A cell of a grid, trimmed; "" when it is empty.
	std::wstring At (const Xlsx::Grid &g, size_t row, size_t col)
		{
		const auto r = g.find (row);
		if (r == g.end ())
			return std::wstring ();
		const auto c = r->second.find (col);
		return c == r->second.end () ? std::wstring () : Trim (c->second);
		}

	/// A count as a person writes it ("40"); "" when the cell holds none.
	std::wstring Whole (const std::wstring &text)
		{
		double v = 0;
		if (!Csv::ParseDouble (text, v) || v < 0)
			return std::wstring ();
		return Csv::Tidy (std::round (v));
		}

	/// A column's place in a header row, or -1.
	int ColOf (const Csv::Row &head, const wchar_t *name)
		{
		for (size_t i = 0; i < head.size (); ++i)
			if (Csv::Trim (head[i]) == name)
				return static_cast<int> (i);
		return -1;
		}

	/// A cell of a row as a number; NaN when it is empty or not one.
	double Num (const Csv::Row &row, int col)
		{
		double v = 0;
		if (col < 0 || static_cast<size_t> (col) >= row.size () || !Csv::ParseDouble (row[static_cast<size_t> (col)], v))
			return kNaN;
		return v;
		}

	/// The op_idn of a row, or 0.
	long OpOf (const Csv::Row &row, int col)
		{
		long long id = 0;
		if (col < 0 || static_cast<size_t> (col) >= row.size () || !Csv::ParseLong (row[static_cast<size_t> (col)], id) || id <= 0)
			return 0;
		return static_cast<long> (id);
		}

	// ---- Dates: days since 1970-01-01 to and from the calendar (the proleptic
	// Gregorian one, which is Excel's after 1900-02-28).

	long long DaysFromCivil (long long y, unsigned m, unsigned d)
		{
		y -= m <= 2 ? 1 : 0;
		const long long era = (y >= 0 ? y : y - 399) / 400;
		const unsigned yoe = static_cast<unsigned> (y - era * 400);
		const unsigned doy = (153 * (m > 2 ? m - 3 : m + 9) + 2) / 5 + d - 1;
		const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
		return era * 146097 + static_cast<long long> (doe) - 719468;
		}

	void CivilFromDays (long long z, long long &y, unsigned &m, unsigned &d)
		{
		z += 719468;
		const long long era = (z >= 0 ? z : z - 146096) / 146097;
		const unsigned doe = static_cast<unsigned> (z - era * 146097);
		const unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
		const unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
		const unsigned mp = (5 * doy + 2) / 153;
		d = doy - (153 * mp + 2) / 5 + 1;
		m = mp < 10 ? mp + 3 : mp - 9;
		y = static_cast<long long> (yoe) + era * 400 + (m <= 2 ? 1 : 0);
		}

	/// Excel's day 0 (1899-12-30) in days since 1970-01-01.
	const long long kExcelEpoch = -25569;

	/// Money to two places, "" for NaN.
	std::wstring Money (double v)
		{
		if (std::isnan (v))
			return std::wstring ();
		wchar_t buf[48];
		swprintf (buf, 48, L"%.2f", std::round (v * 100.0) / 100.0 + 0.0);
		return buf;
		}

	/// A share as a person reads it: "+1.9%", "-5.1%".
	std::wstring Pct (double part, double of)
		{
		if (std::isnan (part) || std::isnan (of) || of <= 0)
			return std::wstring ();
		wchar_t buf[48];
		const double p = std::round (part / of * 1000.0) / 10.0 + 0.0;
		swprintf (buf, 48, L"%+.1f%%", p);
		return buf;
		}

	/// A saving in seconds: "1:21:27", or "-0:00:30" when time was added.
	std::wstring Saved (double seconds)
		{
		if (std::isnan (seconds))
			return std::wstring ();
		const double r = std::round (seconds);
		return (r < 0 ? L"-" : L"") + History::Hms (std::fabs (r));
		}

	/// A measured (or Mastercam's) time against the estimate, in words, after
	/// "the machine was": "0:29:20 slower than the estimate of 25:10:40 (+1.9%)".
	/// "" when either is not known.
	std::wstring Against (double actual, double estimate)
		{
		if (std::isnan (actual) || std::isnan (estimate))
			return std::wstring ();
		const double d = std::round (actual) - std::round (estimate);
		if (d == 0)
			return L"the same as the estimate of " + History::Hms (estimate);
		return History::Hms (std::fabs (d)) + (d > 0 ? L" slower" : L" faster") + L" than the estimate of "
			   + History::Hms (estimate) + L" (" + Pct (d, estimate) + L")";
		}

	/// The key one measurement is known by: its date (when one was typed) and its
	/// figures, as the tool writes them.
	std::wstring MeasurementKey (const History::Record &r)
		{
		const std::wstring date = r.Get (L"date typed") == L"no" ? std::wstring (L"-") : r.Get (L"measured");
		return date + L"|" + History::Hms (History::Seconds (r.Get (L"cycle"))) + L"|" + r.Get (L"parts") + L"|"
			   + r.Get (L"inserts") + L"|" + r.Get (L"note");
		}

	std::wstring WhyKey (const History::Record &r)
		{
		return r.Get (L"op") + L"|" + r.Get (L"column") + L"|" + r.Get (L"change") + L"|" + r.Get (L"why");
		}
	}

namespace History
	{
	// ---- Records and the file ---------------------------------------------------

	std::wstring Record::Get (const std::wstring &name) const
		{
		for (const auto &f : fields)
			if (f.first == name)
				return f.second;
		return std::wstring ();
		}

	bool Record::Has (const std::wstring &name) const
		{
		for (const auto &f : fields)
			if (f.first == name)
				return true;
		return false;
		}

	Record &Record::Set (const std::wstring &name, const std::wstring &value)
		{
		for (auto it = fields.begin (); it != fields.end (); ++it)
			if (it->first == name)
				{
				fields.erase (it);
				break;
				}
		if (!value.empty ())
			fields.push_back ({ name, value });
		return *this;
		}

	std::filesystem::path PathFor (const std::filesystem::path &part)
		{
		return part.parent_path () / (part.stem ().wstring () + L".pthistory");
		}

	std::wstring Line (const Record &r)
		{
		std::wstring o = OneField (r.when) + kSep + OneField (r.kind);
		for (const auto &f : r.fields)
			{
			const std::wstring v = OneField (f.second);
			if (!v.empty ())
				o += kSep + OneField (f.first) + L": " + v;
			}
		return o;
		}

	bool Parse (const std::wstring &line, Record &r)
		{
		std::wstring l = line;
		if (!l.empty () && l[0] == L'﻿')
			l.erase (0, 1);
		l = Trim (l);
		if (l.empty () || l[0] == L'#')
			return false;
		// Split on the bar alone, so a line edited by hand ("a|b") still reads.
		std::vector<std::wstring> parts;
		for (size_t from = 0;;)
			{
			const size_t bar = l.find (L'|', from);
			parts.push_back (Trim (l.substr (from, bar == std::wstring::npos ? std::wstring::npos : bar - from)));
			if (bar == std::wstring::npos)
				break;
			from = bar + 1;
			}
		if (parts.size () < 2 || parts[1].empty ())
			return false;
		r = Record ();
		r.when = parts[0];
		r.kind = Lower (parts[1]);
		for (size_t k = 2; k < parts.size (); ++k)
			{
			// "name: value" - a value may hold ": " itself (a note), a name never does.
			const size_t colon = parts[k].find (L": ");
			if (colon == std::wstring::npos)
				r.fields.push_back ({ Lower (parts[k]), std::wstring () });
			else
				r.fields.push_back ({ Lower (Trim (parts[k].substr (0, colon))), Trim (parts[k].substr (colon + 2)) });
			}
		return true;
		}

	std::vector<Record> Read (const std::filesystem::path &file)
		{
		std::vector<Record> out;
		std::ifstream in (file, std::ios::binary);
		if (!in)
			return out;
		const std::string bytes ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		const std::wstring text = Csv::FromUtf8 (bytes.compare (0, 3, "\xEF\xBB\xBF") == 0 ? bytes.substr (3) : bytes);
		for (size_t at = 0; at < text.size ();)
			{
			size_t end = text.find (L'\n', at);
			if (end == std::wstring::npos)
				end = text.size ();
			Record r;
			if (Parse (text.substr (at, end - at), r))
				out.push_back (r);
			at = end + 1;
			}
		return out;
		}

	bool Append (const std::filesystem::path &file, const std::vector<Record> &records)
		{
		if (records.empty ())
			return true;
		std::error_code ec;
		const bool fresh = !std::filesystem::exists (file, ec) || std::filesystem::file_size (file, ec) == 0;
		std::wstring text;
		if (fresh)
			text = L"# Parameter Table Tool - the improvement history of this part: every dump, load, undo and\r\n"
				   L"# regeneration, what was measured on the machine, and the reasons typed for the changes.\r\n"
				   L"# One record per line, oldest first. The tool only adds lines; the dump shows them on the\r\n"
				   L"# workbook's History sheet. Delete this file to start a new history.\r\n";
		else
			{
			// A file edited by hand may not end with a line break: one first, so the
			// new record is not glued to the last.
			std::ifstream in (file, std::ios::binary);
			in.seekg (-1, std::ios::end);
			char last = '\n';
			if (in && in.get (last) && last != '\n')
				text = L"\r\n";
			}
		for (const Record &r : records)
			text += Line (r) + L"\r\n";
		std::ofstream out (file, std::ios::binary | std::ios::app);
		if (!out)
			return false;
		const std::string bytes = Csv::ToUtf8Bom (text);
		// A BOM only at the top of a new file - further down it would be a stray character.
		out.write (fresh ? bytes.data () : bytes.data () + 3, static_cast<std::streamsize> (fresh ? bytes.size () : bytes.size () - 3));
		return static_cast<bool> (out);
		}

	bool Book::Add (const std::vector<Record> &more)
		{
		if (more.empty ())
			return true;
		if (!Append (file, more))
			return false;
		records.insert (records.end (), more.begin (), more.end ());
		return true;
		}

	Book Open (const std::filesystem::path &part)
		{
		Book b;
		b.file = PathFor (part);
		b.records = Read (b.file);
		return b;
		}

	// ---- Times, dates, figures ----------------------------------------------------

	double Seconds (const std::wstring &text)
		{
		const std::wstring t = Trim (text);
		if (t.empty ())
			return kNaN;
		double v = 0;
		if (t.find (L':') == std::wstring::npos)
			{
			if (!Csv::ParseDouble (t, v) || v < 0)
				return kNaN;
			return v < 1 ? v * 86400.0 : v;
			}
		std::vector<double> parts;
		for (size_t from = 0;;)
			{
			const size_t colon = t.find (L':', from);
			if (!Csv::ParseDouble (t.substr (from, colon == std::wstring::npos ? std::wstring::npos : colon - from), v) || v < 0)
				return kNaN;
			parts.push_back (v);
			if (colon == std::wstring::npos)
				break;
			from = colon + 1;
			}
		if (parts.size () == 2)
			return parts[0] * 60 + parts[1];
		if (parts.size () == 3)
			return parts[0] * 3600 + parts[1] * 60 + parts[2];
		return kNaN;
		}

	std::wstring Hms (double seconds)
		{
		if (std::isnan (seconds) || seconds < 0)
			return std::wstring ();
		const long long s = std::llround (seconds);
		wchar_t buf[48];
		swprintf (buf, 48, L"%lld:%02lld:%02lld", s / 3600, (s / 60) % 60, s % 60);
		return buf;
		}

	std::wstring Change (double seconds)
		{
		if (std::isnan (seconds))
			return std::wstring ();
		const double r = std::round (seconds);
		return (r < 0 ? L"-" : L"+") + Hms (std::fabs (r));
		}

	std::wstring Now ()
		{
		const std::time_t now = std::time (nullptr);
		std::tm t {};
		localtime_s (&t, &now);
		wchar_t buf[32] = L"";
		std::wcsftime (buf, 32, L"%Y-%m-%d %H:%M", &t);
		return buf;
		}

	std::wstring DateOfSerial (double serial)
		{
		// From 1900-03-01 (serial 61), where Excel's calendar is the true one, to 9999.
		if (std::isnan (serial) || serial < 61 || serial >= 2958466)
			return std::wstring ();
		long long y = 0;
		unsigned m = 0, d = 0;
		CivilFromDays (static_cast<long long> (std::floor (serial)) + kExcelEpoch, y, m, d);
		wchar_t buf[32];
		swprintf (buf, 32, L"%04lld-%02u-%02u", y, m, d);
		return buf;
		}

	double SerialOfDate (const std::wstring &date)
		{
		const std::wstring t = Trim (date);
		unsigned y = 0, m = 0, d = 0;
		wchar_t s1 = 0, s2 = 0;
		if (swscanf_s (t.c_str (), L"%4u%c%2u%c%2u", &y, &s1, 1, &m, &s2, 1, &d) != 5 || s1 != s2 || (s1 != L'-' && s1 != L'/')
			|| y < 1901 || y > 9999 || m < 1 || m > 12 || d < 1)
			return kNaN;
		const long long days = DaysFromCivil (y, m, d);
		// The day must exist (2026-02-30 is not one).
		long long yy = 0;
		unsigned mm = 0, dd = 0;
		CivilFromDays (days, yy, mm, dd);
		if (yy != static_cast<long long> (y) || mm != m || dd != d)
			return kNaN;
		return static_cast<double> (days - kExcelEpoch);
		}

	std::wstring Figure (double v)
		{
		if (std::isnan (v))
			return std::wstring ();
		return Csv::Tidy (std::round (v * 100.0) / 100.0 + 0.0);
		}

	std::pair<double, double> Span (const std::wstring &text, bool time)
		{
		auto one = [time] (const std::wstring &t)
			{
			double v = 0;
			if (time)
				return Seconds (t);
			return Csv::ParseDouble (Trim (t), v) ? v : kNaN;
			};
		const size_t a = text.find (kArrow);
		if (a == std::wstring::npos)
			{
			const double v = one (text);
			return { v, v };
			}
		return { one (text.substr (0, a)), one (text.substr (a + kArrow.size ())) };
		}

	// ---- What a workbook says -----------------------------------------------------

	bool ReadActual (const Xlsx::Grid &summary, Actual &a)
		{
		a = Actual ();
		auto valueOf = [&summary] (const wchar_t *label)
			{
			const std::wstring want = Lower (label);
			for (const auto &row : summary)
				{
				const auto b = row.second.find (1);
				if (b != row.second.end () && Lower (Trim (b->second)) == want)
					return At (summary, row.first, 2);
				}
			return std::wstring ();
			};
		const double secs = Seconds (valueOf (Summary::Machine::kCycle));
		if (std::isnan (secs) || secs <= 0)
			return false;
		a.cycle = Hms (secs);
		// The date: a real date is a number to Excel; a text that is a date is taken
		// too, and any other text is kept as typed.
		const std::wstring date = valueOf (Summary::Machine::kDate);
		double serial = 0;
		if (Csv::ParseDouble (date, serial))
			a.measured = DateOfSerial (serial);
		else if (!std::isnan (SerialOfDate (date)))
			a.measured = DateOfSerial (SerialOfDate (date));
		else
			a.measured = OneField (date);
		a.parts = Whole (valueOf (Summary::Machine::kParts));
		a.inserts = Whole (valueOf (Summary::Machine::kInserts));
		a.note = OneField (valueOf (Summary::Machine::kNote));
		a.estimate = Hms (Seconds (valueOf (Summary::Machine::kEstimate)));
		return true;
		}

	std::vector<Record> Measured (const std::filesystem::path &workbook, const std::vector<Record> &have,
								  const std::wstring &now)
		{
		Xlsx::Grid summary;
		std::wstring why;
		Actual a;
		if (!Xlsx::ReadGrid (workbook, L"Summary", summary, why) || !ReadActual (summary, a))
			return {};
		Record r;
		r.when = now;
		r.kind = L"actual";
		r.Set (L"measured", a.measured.empty () ? now.substr (0, 10) : a.measured);
		if (a.measured.empty ())
			r.Set (L"date typed", L"no");
		r.Set (L"cycle", a.cycle).Set (L"parts", a.parts).Set (L"inserts", a.inserts).Set (L"estimate", a.estimate);
		r.Set (L"note", a.note);
		// Each measurement once: the next dump shows the newest again, and a workbook
		// is read at every dump and load.
		const std::wstring key = MeasurementKey (r);
		for (const Record &h : have)
			if (h.kind == L"actual" && MeasurementKey (h) == key)
				return {};
		return { r };
		}

	std::vector<Record> Whys (const std::filesystem::path &workbook, const std::vector<Record> &have,
							  const std::wstring &now)
		{
		std::vector<Record> out;
		Xlsx::Grid rep;
		std::wstring why;
		if (!Xlsx::ReadGrid (workbook, L"Change report", rep, why))
			return out;
		// The heading row of the changes: "Why (type the reason)" and "key" (op|column,
		// a hidden column) in it; "As dumped" and "Now" the values.
		size_t head = 0, keyCol = 0, whyCol = 0, wasCol = 4, nowCol = 5;
		for (const auto &row : rep)
			{
			size_t k = 0, w = 0, was = 0, nw = 0;
			for (const auto &cell : row.second)
				{
				const std::wstring h = Lower (Trim (cell.second));
				if (h == L"key") k = cell.first + 1;
				else if (h.rfind (L"why", 0) == 0) w = cell.first + 1;
				else if (h == L"as dumped") was = cell.first + 1;
				else if (h == L"now") nw = cell.first + 1;
				}
			if (k > 0 && w > 0)
				{
				head = row.first;
				keyCol = k - 1;
				whyCol = w - 1;
				if (was > 0) wasCol = was - 1;
				if (nw > 0) nowCol = nw - 1;
				break;
				}
			}
		if (head == 0)
			return out;
		std::set<std::wstring> known;
		for (const Record &h : have)
			if (h.kind == L"why")
				known.insert (WhyKey (h));
		for (const auto &row : rep)
			{
			if (row.first <= head)
				continue;
			const std::wstring key = At (rep, row.first, keyCol), text = At (rep, row.first, whyCol);
			const size_t bar = key.find (L'|');
			if (text.empty () || bar == std::wstring::npos)
				continue;
			Record r;
			r.when = now;
			r.kind = L"why";
			r.Set (L"op", key.substr (0, bar)).Set (L"column", key.substr (bar + 1));
			r.Set (L"change", At (rep, row.first, wasCol) + kArrow + At (rep, row.first, nowCol)).Set (L"why", OneField (text));
			if (known.insert (WhyKey (r)).second)
				out.push_back (r);
			}
		return out;
		}

	const Record *NewestActual (const std::vector<Record> &records)
		{
		const Record *best = nullptr;
		for (const Record &r : records)
			if (r.kind == L"actual" && (best == nullptr || r.Get (L"measured") >= best->Get (L"measured")))
				best = &r;
		return best;
		}

	const Record *Last (const std::vector<Record> &records, const std::wstring &kind)
		{
		for (auto it = records.rbegin (); it != records.rend (); ++it)
			if (it->kind == kind)
				return &*it;
		return nullptr;
		}

	std::map<long, double> Estimates (const Record &load)
		{
		std::map<long, double> out;
		const std::wstring list = load.Get (L"estimates");
		for (size_t from = 0; from < list.size ();)
			{
			size_t comma = list.find (L',', from);
			if (comma == std::wstring::npos)
				comma = list.size ();
			const std::wstring item = Trim (list.substr (from, comma - from));
			from = comma + 1;
			const size_t space = item.find (L' ');
			long long op = 0;
			if (!Csv::ParseLong (item.substr (0, space), op) || op <= 0)
				continue;
			out[static_cast<long> (op)] = space == std::wstring::npos ? kNaN : Seconds (item.substr (space + 1));
			}
		return out;
		}

	double InsertCost (const std::vector<Csv::Row> &main, const std::vector<Csv::Row> &dumped,
					   const Xlsx::Grid &tools, const std::set<long> &now)
		{
		if (main.empty ())
			return kNaN;
		const int idCol = ColOf (main[0], L"op_idn"), toolCol = ColOf (main[0], L"tool");
		const int flipsCol = ColOf (main[0], L"flips_part"), copiesCol = ColOf (main[0], L"xf_copies");
		if (idCol < 0 || toolCol < 0 || flipsCol < 0)
			return kNaN;
		std::map<long, double> was;
		if (!dumped.empty ())
			{
			const int dId = ColOf (dumped[0], L"op_idn"), dFlips = ColOf (dumped[0], L"flips_part");
			for (size_t k = 1; k < dumped.size (); ++k)
				{
				const long op = OpOf (dumped[k], dId);
				const double v = Num (dumped[k], dFlips);
				if (op != 0 && !std::isnan (v))
					was[op] = v;
				}
			}
		// Each tool's flips per part, a transform's copies counted - as its Tools row sums them.
		std::map<std::wstring, double> toolFlips;
		for (size_t k = 1; k < main.size (); ++k)
			{
			const long op = OpOf (main[k], idCol);
			const std::wstring tool = static_cast<size_t> (toolCol) < main[k].size () ? Csv::Trim (main[k][static_cast<size_t> (toolCol)])
																					  : std::wstring ();
			if (op == 0 || tool.empty ())
				continue;
			double f = Num (main[k], flipsCol);
			if (!dumped.empty () && now.count (op) == 0)
				{
				const auto at = was.find (op);
				f = at != was.end () ? at->second : kNaN;
				}
			if (std::isnan (f))
				continue;
			const double copies = Num (main[k], copiesCol);
			toolFlips[tool] += f * (1 + (std::isnan (copies) ? 0 : copies));
			}
		// The Tools page: each tool's insert, by its heading.
		size_t insCol = 3;
		const auto heads = tools.find (1);
		if (heads != tools.end ())
			for (const auto &h : heads->second)
				if (Lower (Trim (h.second)) == L"insert")
					insCol = h.first;
		std::map<std::wstring, double> insertFlips;
		for (size_t r = 2; !At (tools, r, 0).empty (); ++r)
			{
			const std::wstring name = At (tools, r, insCol);
			const auto f = toolFlips.find (At (tools, r, 0));
			if (!name.empty ())
				insertFlips[name] += f != toolFlips.end () ? f->second : 0;
			}
		// The inserts table: per insert, flips / edges x cost.
		for (const auto &row : tools)
			{
			if (At (tools, row.first, 0) != L"Inserts")
				continue;
			size_t nameCol = 1, edgesCol = 3, costCol = 6, partsCol = 8;
			for (const auto &h : row.second)
				{
				const std::wstring head = Lower (Trim (h.second));
				if (head == L"insert") nameCol = h.first;
				else if (head.rfind (L"edges", 0) == 0) edgesCol = h.first;
				else if (head == L"cost per insert") costCol = h.first;
				else if (head.rfind (L"parts per edge", 0) == 0) partsCol = h.first;
				}
			double total = 0;
			bool any = false;
			for (size_t r = row.first + 1; tools.count (r); ++r)
				{
				const std::wstring name = At (tools, r, nameCol);
				double edges = 0, cost = 0, ppe = 0;
				if (name.empty () || !Csv::ParseDouble (At (tools, r, costCol), cost))
					continue;
				any = true;
				if (!Csv::ParseDouble (At (tools, r, edgesCol), edges) || edges <= 0)
					continue;
				const auto f = insertFlips.find (name);
				double flips = f != insertFlips.end () ? f->second : 0;
				// An insert that outlasts a part: 1 / parts per edge, where the ops flip it less than once.
				if (Csv::ParseDouble (At (tools, r, partsCol), ppe) && ppe > 0 && flips < 1)
					flips = 1.0 / ppe;
				total += flips / edges * cost;
				}
			return any ? total : kNaN;
			}
		return kNaN;
		}

	Xlsx::Grid ToolsGrid (const Xlsx::Sheet &s)
		{
		Xlsx::Grid g;
		if (s.tools.empty ())
			return g;
		g[1][0] = L"Tool";
		g[1][1] = L"Name";
		g[1][2] = L"Used by";
		for (size_t e = 0; e < s.toolExtraHeads.size (); ++e)
			g[1][3 + e] = s.toolExtraHeads[e];
		for (size_t k = 0; k < s.tools.size (); ++k)
			{
			const Xlsx::Sheet::ToolRow &t = s.tools[k];
			std::map<size_t, std::wstring> &row = g[k + 2];
			row[0] = t.number;
			row[1] = t.name;
			row[2] = t.usedBy;
			for (size_t e = 0; e < t.extra.size (); ++e)
				if (!t.extra[e].text.empty ())
					row[3 + e] = t.extra[e].text;
			}
		for (size_t k = 0; k < s.toolsAfter.size (); ++k)
			{
			if (s.toolsAfter[k].empty ())
				continue;
			std::map<size_t, std::wstring> &row = g[s.tools.size () + 3 + k];
			for (size_t c = 0; c < s.toolsAfter[k].size (); ++c)
				if (!s.toolsAfter[k][c].text.empty ())
					row[c] = s.toolsAfter[k][c].text;
			}
		return g;
		}

	// ---- The records the tool writes ----------------------------------------------

	Record DumpRecord (const Xlsx::Sheet &s, const std::wstring &file, bool wholePart,
					   const std::wstring &batch, const std::wstring &when)
		{
		Record r;
		r.when = when;
		r.kind = L"dump";
		const size_t ops = s.rows.empty () ? 0 : s.rows.size () - 1;
		r.Set (L"file", file).Set (L"ops", std::to_wstring (ops)).Set (L"whole part", wholePart ? L"yes" : L"no");
		if (!s.rows.empty ())
			{
			// The part's figures as its Summary sums them: the estimate (Mastercam's own
			// at a dump) and the flips per part.
			const int estCol = ColOf (s.rows[0], L"est_seconds"), flipsCol = ColOf (s.rows[0], L"flips_part");
			double est = 0, flips = 0;
			bool anyEst = false, anyFlips = false;
			for (size_t k = 1; k < s.rows.size (); ++k)
				{
				const double e = Num (s.rows[k], estCol), f = Num (s.rows[k], flipsCol);
				if (!std::isnan (e))
					{
					est += e;
					anyEst = true;
					}
				if (!std::isnan (f))
					{
					flips += f;
					anyFlips = true;
					}
				}
			if (anyEst)
				r.Set (L"cycle", Hms (est));
			if (anyFlips)
				r.Set (L"flips", Figure (flips));
			r.Set (L"insert cost", Figure (InsertCost (s.rows, {}, ToolsGrid (s), {})));
			}
		r.Set (L"batch", batch);
		return r;
		}

	Record LoadRecord (const LoadFigures &f, const std::wstring &when)
		{
		Record r;
		r.when = when;
		r.kind = L"load";
		r.Set (L"file", f.files).Set (L"changes", std::to_wstring (f.changes)).Set (L"ops", std::to_wstring (f.ops));
		if (f.failed > 0)
			r.Set (L"failed", std::to_wstring (f.failed));
		auto span = [] (double a, double b, bool time) -> std::wstring
			{
			if (std::isnan (a) || std::isnan (b))
				return std::wstring ();
			return time ? Hms (a) + kArrow + Hms (b) : Figure (a) + kArrow + Figure (b);
			};
		r.Set (L"cycle", span (f.cycleWas, f.cycleNow, true));
		r.Set (L"flips", span (f.flipsWas, f.flipsNow, false));
		r.Set (L"insert cost", span (f.costWas, f.costNow, false));
		r.Set (L"batch", f.batch);
		std::wstring list;
		for (const auto &e : f.estimates)
			list += (list.empty () ? L"" : L", ") + std::to_wstring (e.first) + L" "
					+ (std::isnan (e.second) ? std::wstring (L"-") : Hms (e.second));
		r.Set (L"estimates", list);
		return r;
		}

	std::vector<Record> RegenRecords (const std::vector<Regenerated> &ops, double partEstimate,
									  const std::wstring &when)
		{
		std::vector<Record> out;
		Record all;
		all.when = when;
		all.kind = L"regen";
		double est = 0, mc = 0;
		size_t both = 0, failed = 0;
		for (const Regenerated &o : ops)
			{
			if (o.failed)
				++failed;
			if (!std::isnan (o.estimate) && !std::isnan (o.mastercam))
				{
				est += o.estimate;
				mc += o.mastercam;
				++both;
				}
			}
		all.Set (L"ops", std::to_wstring (ops.size ()));
		if (failed > 0)
			all.Set (L"failed", std::to_wstring (failed));
		if (both > 0)
			{
			all.Set (L"estimate", Hms (est)).Set (L"mastercam", Hms (mc));
			if (both < ops.size ())
				all.Set (L"compared", std::to_wstring (both));
			// The whole part with these ops at Mastercam's own figures.
			if (!std::isnan (partEstimate))
				all.Set (L"part estimate", Hms (partEstimate)).Set (L"part mastercam", Hms (partEstimate - est + mc));
			}
		out.push_back (all);
		for (const Regenerated &o : ops)
			{
			Record r;
			r.when = when;
			r.kind = L"regen";
			r.Set (L"op", std::to_wstring (o.op)).Set (L"estimate", Hms (o.estimate)).Set (L"mastercam", Hms (o.mastercam));
			if (o.failed)
				r.Set (L"failed", L"yes");
			out.push_back (r);
			}
		return out;
		}

	// ---- The History sheet -----------------------------------------------------------

	Xlsx::Sheet::Page Page (const std::vector<Record> &records, const std::wstring &partName,
							const std::wstring &historyFile, const std::wstring &batchNow)
		{
		using Cell = Xlsx::Sheet::FreeCell;
		Xlsx::Sheet::Page pg;
		pg.name = L"History";
		// When | What happened | Cycle time / part | Saved / part | Saved / batch | Saved so far | Flips | Insert cost
		pg.widths = { 17, 58, 21, 12, 17, 14, 13, 17 };
		auto cell = [] (const std::wstring &t, Cell::Look look = Cell::Plain, bool right = false)
			{
			Cell c;
			c.text = t;
			c.look = look;
			c.right = right;
			return c;
			};
		auto figure = [&cell] (const std::wstring &t) { return cell (t, Cell::Plain, true); };
		auto head = [] (const std::wstring &t, bool right)
			{
			Cell c;
			c.text = t;
			c.head = true;
			c.look = Cell::Plain;
			c.right = right;
			return c;
			};
		auto line = [&pg, &cell] (const std::wstring &t, Cell::Look look = Cell::Plain)
			{
			pg.rows.push_back ({ cell (t, look) });
			};

		// ---- THE RECORDS IN GROUPS: a load with the reasons written straight after
		// it, a regeneration with its ops - each group a line, its details under it.
		struct Group
			{
			const Record *main = nullptr;
			std::vector<const Record *> details;
			double saved = kNaN, savedBatch = kNaN, total = kNaN;	//!< seconds per part / per batch, running total
			double flipsSaved = kNaN, costSaved = kNaN;
			std::wstring qty;					//!< the batch quantity the batch figure is for
			std::wstring undone;				//!< a load: when an undo took it back ("" = never)
			bool undoneWhole = false;
			};
		std::vector<Group> groups;
		for (const Record &r : records)
			{
			if (!groups.empty () && r.when == groups.back ().main->when
				&& ((r.kind == L"why" && groups.back ().main->kind == L"load")
					|| (r.kind == L"regen" && r.Has (L"op") && groups.back ().main->kind == L"regen"
						&& !groups.back ().main->Has (L"op"))))
				{
				groups.back ().details.push_back (&r);
				continue;
				}
			Group g;
			g.main = &r;
			groups.push_back (g);
			}

		// ---- WHAT THE LOADS SAVED, in the order they happened: each load's saving
		// per part and per batch, and the running total - an undo of a whole load
		// takes its saving back.
		double total = 0, flipsTotal = 0, costTotal = 0, batchTotal = 0;
		bool anyTime = false, anyFlips = false, anyCost = false;
		int loads = 0, undos = 0;
		std::vector<size_t> active;			// loads not undone, oldest first
		for (size_t i = 0; i < groups.size (); ++i)
			{
			Group &g = groups[i];
			const Record &r = *g.main;
			if (r.kind == L"load")
				{
				++loads;
				g.qty = r.Get (L"batch").empty () ? batchNow : r.Get (L"batch");
				double qty = 0;
				if (!Csv::ParseDouble (g.qty, qty) || qty < 1)
					{
					qty = 1;
					g.qty = L"1";
					}
				const auto cyc = Span (r.Get (L"cycle"), true);
				if (!std::isnan (cyc.first) && !std::isnan (cyc.second))
					{
					g.saved = std::round (cyc.first) - std::round (cyc.second);
					g.savedBatch = g.saved * qty;
					total += g.saved;
					batchTotal += g.savedBatch;
					anyTime = true;
					}
				const auto fl = Span (r.Get (L"flips"), false), co = Span (r.Get (L"insert cost"), false);
				if (!std::isnan (fl.first) && !std::isnan (fl.second))
					{
					g.flipsSaved = fl.first - fl.second;
					flipsTotal += g.flipsSaved;
					anyFlips = true;
					}
				if (!std::isnan (co.first) && !std::isnan (co.second))
					{
					g.costSaved = co.first - co.second;
					costTotal += g.costSaved;
					anyCost = true;
					}
				g.total = total;
				active.push_back (i);
				}
			else if (r.kind == L"undo")
				{
				if (!active.empty ())
					{
					Group &load = groups[active.back ()];
					active.pop_back ();
					load.undone = r.when;
					load.undoneWhole = r.Get (L"complete") == L"yes";
					if (load.undoneWhole)
						{
						++undos;
						g.qty = load.qty;
						if (!std::isnan (load.saved))
							{
							g.saved = -load.saved;
							g.savedBatch = -load.savedBatch;
							total -= load.saved;
							batchTotal -= load.savedBatch;
							}
						if (!std::isnan (load.flipsSaved))
							flipsTotal -= load.flipsSaved;
						if (!std::isnan (load.costSaved))
							costTotal -= load.costSaved;
						}
					}
				g.total = total;
				}
			}

		// ---- THE PAGE: a title, what the loads saved so far, the last check
		// against the machine and against Mastercam, then every record.
		std::wstring file = historyFile;
		const size_t slash = file.find_last_of (L"\\/");
		if (slash != std::wstring::npos)
			file = file.substr (slash + 1);
		line (L"History - " + partName, Cell::Title);
		line (L"Every dump, load, undo and regeneration of this part, what it took on the machine, and the reasons typed for the "
			  L"changes - newest first. Kept in " + (file.empty () ? std::wstring (L"the part's .pthistory file") : file)
			  + L" beside the part.", Cell::Note);
		pg.rows.push_back ({});
		line (L"So far", Cell::Section);
		double qtyNow = 0;
		if (!Csv::ParseDouble (batchNow, qtyNow) || qtyNow < 1)
			qtyNow = 1;
		const std::wstring qtyText = Csv::Tidy (qtyNow);
		if (loads == 0)
			line (L"No loads yet. What a change saves is counted from the first load.");
		else
			{
			// Against the part's cycle time at its first dump of the whole part.
			double first = kNaN;
			for (const Record &r : records)
				if (r.kind == L"dump" && r.Get (L"whole part") != L"no" && !std::isnan (Seconds (r.Get (L"cycle"))))
					{
					first = Seconds (r.Get (L"cycle"));
					break;
					}
			const std::wstring loadsText = std::to_wstring (loads) + (loads == 1 ? L" load" : L" loads")
										   + (undos > 0 ? L" (" + std::to_wstring (undos) + L" undone)" : L"");
			if (anyTime)
				{
				const std::wstring share = !std::isnan (first) && first > 0
											   ? L" - " + Pct (std::fabs (total), first).substr (1) + L" of the first dump's "
													 + Hms (first)
											   : std::wstring ();
				line (loadsText + L". Time " + (total < 0 ? L"added" : L"saved") + L" per part: " + Hms (std::fabs (std::round (total)))
					  + share + L".");
				line (L"Per batch (at each load's batch quantity): " + Hms (std::fabs (std::round (batchTotal)))
					  + (batchTotal < 0 ? L" added." : L" saved.") + (qtyNow > 1 ? L" A batch of " + qtyText + L" now: "
					  + Hms (std::fabs (std::round (total * qtyNow))) + L"." : std::wstring ()));
				}
			else
				line (loadsText + L". The workbooks loaded had no time estimate to compare.");
			if (anyFlips)
				line (L"Insert flips per part: " + Figure (std::fabs (flipsTotal)) + (flipsTotal < 0 ? L" more." : L" fewer."));
			if (anyCost)
				line (L"Insert cost per part: " + Money (std::fabs (costTotal)) + (costTotal < 0 ? L" more" : L" less")
					  + (qtyNow > 1 ? L" - " + Money (std::fabs (costTotal * qtyNow)) + L" a batch of " + qtyText : std::wstring ())
					  + L".");
			}
		if (const Record *a = NewestActual (records))
			{
			const std::wstring against = Against (Seconds (a->Get (L"cycle")), Seconds (a->Get (L"estimate")));
			line (L"Last measured on the machine (" + a->Get (L"measured") + L"): " + a->Get (L"cycle") + L" a part"
				  + (against.empty () ? std::wstring () : L" - " + against) + L".");
			}
		else
			line (L"Nothing measured on the machine yet - type a run's time under \"From the machine\" on the Summary page.");
		for (auto it = groups.rbegin (); it != groups.rend (); ++it)
			if (it->main->kind == L"regen" && !it->main->Has (L"op") && it->main->Has (L"mastercam"))
				{
				const Record &r = *it->main;
				line (L"Last regeneration (" + r.when + L"): Mastercam was "
					  + Against (Seconds (r.Get (L"mastercam")), Seconds (r.Get (L"estimate"))) + L" for the operations regenerated.");
				break;
				}
		pg.rows.push_back ({});

		line (L"Every record, newest first", Cell::Section);
		pg.rows.push_back ({ head (L"When", false), head (L"What happened", false), head (L"Cycle time / part", true),
							 head (L"Saved / part", true), head (L"Saved / batch (qty)", true), head (L"Saved so far", true),
							 head (L"Flips / part", true), head (L"Insert cost / part", true) });
		pg.titleRow = pg.rows.size ();
		if (groups.empty ())
			line (L"Nothing yet.", Cell::Note);

		// A detail under its line: indented, in the quieter note style.
		auto detail = [&pg, &cell] (const std::wstring &t, const std::wstring &c = std::wstring ())
			{
			std::vector<Cell> row = { cell (L""), cell (L"    " + t, Cell::Note) };
			if (!c.empty ())
				row.push_back (cell (c, Cell::Note, true));
			pg.rows.push_back (row);
			};
		auto plural = [] (const std::wstring &n, const wchar_t *one, const wchar_t *many)
			{
			return n + L" " + (n == L"1" ? one : many);
			};
		for (auto it = groups.rbegin (); it != groups.rend (); ++it)
			{
			const Group &g = *it;
			const Record &r = *g.main;
			std::vector<Cell> row (8, cell (L""));		// Plain: an empty one is no cell at all
			row[0] = cell (r.when);
			std::vector<std::wstring> more;			// lines under it
			if (r.kind == L"dump")
				{
				row[1] = cell (L"Dumped " + plural (r.Get (L"ops"), L"operation", L"operations")
							   + (r.Get (L"whole part") == L"no" ? L" (a selection)" : L" (the whole part)"));
				row[2] = figure (r.Get (L"cycle"));
				row[6] = figure (r.Get (L"flips"));
				double cost = 0;
				row[7] = figure (Csv::ParseDouble (r.Get (L"insert cost"), cost) ? Money (cost) : std::wstring ());
				more.push_back (L"to " + r.Get (L"file") + (r.Get (L"batch").empty () ? L"" : L" - batch of " + r.Get (L"batch")));
				}
			else if (r.kind == L"load")
				{
				row[1] = cell (L"Loaded " + plural (r.Get (L"changes"), L"change", L"changes") + L" on "
							   + plural (r.Get (L"ops"), L"operation", L"operations")
							   + (r.Get (L"failed").empty () ? L"" : L" (" + r.Get (L"failed") + L" failed)"));
				row[2] = figure (r.Get (L"cycle"));
				row[3] = figure (Saved (g.saved));
				row[4] = figure (std::isnan (g.savedBatch) ? std::wstring () : Saved (g.savedBatch) + L" (" + g.qty + L")");
				row[5] = figure (Saved (g.total));
				row[6] = figure (r.Get (L"flips"));
				const auto co = Span (r.Get (L"insert cost"), false);
				row[7] = figure (std::isnan (co.first) || std::isnan (co.second) ? std::wstring ()
																				 : Money (co.first) + kArrow + Money (co.second));
				more.push_back (L"from " + r.Get (L"file"));
				if (!g.undone.empty ())
					more.push_back (g.undoneWhole ? L"undone " + g.undone + L" - its saving is taken back there"
												  : L"partly undone " + g.undone + L" - its saving still counts");
				}
			else if (r.kind == L"undo")
				{
				row[1] = cell (L"Undid the last load: " + plural (r.Get (L"restored"), L"value", L"values") + L" put back");
				row[3] = figure (Saved (g.saved));
				row[4] = figure (std::isnan (g.savedBatch) ? std::wstring () : Saved (g.savedBatch) + L" (" + g.qty + L")");
				row[5] = figure (Saved (g.total));
				if (!r.Get (L"load").empty ())
					more.push_back (L"the load of " + r.Get (L"load"));
				if (r.Get (L"complete") != L"yes")
					more.push_back (L"only part of it - the load's saving still counts");
				}
			else if (r.kind == L"regen" && !r.Has (L"op"))
				{
				row[1] = cell (L"Regenerated " + plural (r.Get (L"ops"), L"operation", L"operations")
							   + (r.Get (L"failed").empty () ? L"" : L" (" + r.Get (L"failed") + L" failed)"));
				row[2] = figure (r.Get (L"part mastercam").empty () ? std::wstring () : r.Get (L"part mastercam") + L" (Mastercam)");
				const std::wstring against = Against (Seconds (r.Get (L"mastercam")), Seconds (r.Get (L"estimate")));
				if (!against.empty ())
					more.push_back (L"Mastercam's time for them, " + r.Get (L"mastercam") + L", was " + against
									+ (r.Get (L"compared").empty () ? L"" : L" (" + r.Get (L"compared") + L" ops compared)"));
				if (!r.Get (L"part estimate").empty ())
					more.push_back (L"the whole part: estimate " + r.Get (L"part estimate") + L", with these ops at Mastercam's time "
									+ r.Get (L"part mastercam"));
				}
			else if (r.kind == L"actual")
				{
				row[1] = cell (L"Measured on the machine on " + r.Get (L"measured")
							   + (r.Get (L"parts").empty () ? std::wstring () : L": " + plural (r.Get (L"parts"), L"part", L"parts")));
				row[2] = figure (r.Get (L"cycle"));
				double n = 0, ins = 0;
				if (Csv::ParseDouble (r.Get (L"inserts"), ins))
					{
					const bool per = Csv::ParseDouble (r.Get (L"parts"), n) && n > 0;
					more.push_back (plural (r.Get (L"inserts"), L"insert", L"inserts") + L" used"
									+ (per ? L" - " + Figure (ins / n) + L" a part" : std::wstring ()));
					}
				const std::wstring against = Against (Seconds (r.Get (L"cycle")), Seconds (r.Get (L"estimate")));
				if (!against.empty ())
					more.push_back (L"the machine was " + against);
				if (!r.Get (L"note").empty ())
					more.push_back (L"note: " + r.Get (L"note"));
				}
			else if (r.kind == L"why")
				{
				row[1] = cell (L"Why op " + r.Get (L"op") + L" " + r.Get (L"column") + L" changed");
				more.push_back (r.Get (L"change") + L": " + r.Get (L"why"));
				}
			else
				{
				// A kind this version does not know (a later version's, or typed by hand).
				row[1] = cell (r.kind);
				std::wstring all;
				for (const auto &f : r.fields)
					all += (all.empty () ? L"" : L", ") + f.first + (f.second.empty () ? L"" : L": " + f.second);
				if (!all.empty ())
					more.push_back (all);
				}
			pg.rows.push_back (row);
			for (const std::wstring &m : more)
				detail (m);
			for (const Record *d : g.details)
				{
				if (d->kind == L"why")
					detail (L"why op " + d->Get (L"op") + L" " + d->Get (L"column") + L" (" + d->Get (L"change") + L"): "
							+ d->Get (L"why"));
				else
					{
					const std::wstring against = Against (Seconds (d->Get (L"mastercam")), Seconds (d->Get (L"estimate")));
					detail (L"op " + d->Get (L"op") + (d->Get (L"failed") == L"yes" ? std::wstring (L": the regeneration FAILED")
													   : against.empty () ? L": Mastercam " + d->Get (L"mastercam")
													   : L": Mastercam was " + against),
							d->Get (L"mastercam"));
					}
				}
			}
		return pg;
		}
	}
