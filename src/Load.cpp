#include "stdafx.h"
#include "MastercamSdk.h"
#include "Load.h"
#include "LatheFields.h"
#include "Coolant.h"
#include "Plan.h"
#include "Util.h"
#include "Csv.h"
#include "Xlsx.h"
#include "Preview.h"
#include "Settings.h"

#include <algorithm>
#include <fstream>
#include <iterator>
#include <map>
#include <sstream>

namespace
	{
	/// What one CSV file would do, with the table it was matched to.
	struct FilePlan
		{
		std::filesystem::path file;
		const Lathe::Table *table = nullptr;
		Plan::Result result;
		std::map<long, size_t> opRow;	//!< each operation's row in the sheet
		};

	bool ReadFile (const std::filesystem::path &file, std::wstring &text)
		{
		std::ifstream in (file, std::ios::binary);
		if (!in)
			return false;
		const std::string bytes ((std::istreambuf_iterator<char> (in)),
								 std::istreambuf_iterator<char> ());
		text = Csv::FromUtf8 (bytes);
		return true;
		}

	/// The operations of one kind as they stand right now, as CSV-comparable text.
	std::vector<Plan::Current> Snapshot (const Lathe::Table &t)
		{
		std::vector<Plan::Current> now;

		Cnc::Tool::TpPartOpList &opList = TpMainOpMgr.GetMainOpList ();
		const INT_PTR n = opList.GetSize ();
		for (INT_PTR i = 0; i < n; ++i)
			{
			operation *pOp = opList.GetAt (i);
			if (pOp == nullptr || Lathe::TableFor (static_cast<long> (pOp->opcode)) != &t)
				continue;
			void *prm = Lathe::PrmFor (*pOp, t.opcode);
			if (prm == nullptr)
				continue;

			Plan::Current c;
			c.op = pOp->op_idn;
			c.type = t.schema.type;
			for (const Lathe::Binding &b : t.bindings)
				c.values.push_back (Lathe::Read (b, pOp, prm));
			now.push_back (c);
			}
		return now;
		}

	/// A value on one log line: a manual entry's line breaks shown as the
	/// preview shows them, so one change stays one line.
	std::wstring OneLine (const std::wstring &v)
		{
		std::wstring o;
		for (wchar_t c : v)
			{
			if (c == L'\r')
				continue;
			o += c == L'\n' ? std::wstring (L" \u21B5 ") : std::wstring (1, c);
			}
		return o;
		}

	/// The newest dumped sheet of this part, or "" - what a load is almost
	/// always about to pick.
	std::filesystem::path NewestDump (const std::filesystem::path &part)
		{
		const std::wstring prefix = part.stem ().wstring () + L"_lathe_params_";
		std::filesystem::path best;
		std::filesystem::file_time_type bestTime;
		std::error_code ec;
		for (const auto &e : std::filesystem::directory_iterator (part.parent_path (), ec))
			{
			const std::wstring name = e.path ().filename ().wstring ();
			std::wstring ext = e.path ().extension ().wstring ();
			for (wchar_t &c : ext)
				c = static_cast<wchar_t> (towlower (c));
			if (name.compare (0, prefix.size (), prefix) != 0
				|| (ext != L".xlsx" && ext != L".xlsm" && ext != L".csv"))
				continue;
			const auto t = e.last_write_time (ec);
			if (best.empty () || t > bestTime)
				{
				best = e.path ();
				bestTime = t;
				}
			}
		return best;
		}

	/// The table for a type label ("ROUGH"), or nullptr for one this tool does
	/// not know.
	const Lathe::Table *TableForType (const std::wstring &type)
		{
		for (const Lathe::Table &t : Lathe::AllTables ())
			if (t.schema.type == type)
				return &t;
		return nullptr;
		}

	/// This kind's schema as the SHEET sees it: the same columns, plus a note of
	/// every sheet column that belongs to a different kind - so a value typed
	/// into one is refused instead of ignored.
	Plan::Schema SheetSchema (const Lathe::Table &t)
		{
		Plan::Schema s = t.schema;
		for (const std::wstring &name : Lathe::SheetColumns ())
			if (Lathe::IndexOf (t, name) < 0)
				s.foreign.push_back (name);
		return s;
		}
	}

namespace Load
	{
	int Run ()
		{
		Coolant::Reset ();
		const std::filesystem::path part = Util::PartFile ();
		if (part.empty ())
			{
			Util::Say (L"Save the part first - there is no open part to load "
					   L"changes into.");
			return 0;
			}

		// ---- Which files.
		CFileDialog dlg (TRUE, L"xlsx", nullptr,
						 OFN_ALLOWMULTISELECT | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY,
						 L"Dumped sheets (*.xlsx;*.xlsm;*.csv)|*.xlsx;*.xlsm;*.csv|All files (*.*)|*.*||",
						 CWnd::FromHandle (get_MainFrame ()->GetSafeHwnd ()));

		std::vector<wchar_t> buffer (32768, 0);
		// Where to start: the sheet this part was last dumped to, wherever it was
		// saved; else the newest beside the part.
		std::wstring newest = Settings::LastDumpFor (part.wstring ());
		{
		std::error_code ec;
		if (newest.empty () || !std::filesystem::exists (newest, ec))
			newest = NewestDump (part).wstring ();
		}
		if (!newest.empty () && newest.size () < buffer.size ())
			std::copy (newest.begin (), newest.end (), buffer.begin ());
		dlg.m_ofn.lpstrFile = buffer.data ();
		dlg.m_ofn.nMaxFile = static_cast<DWORD> (buffer.size ());
		const std::wstring startDir = part.parent_path ().wstring ();
		dlg.m_ofn.lpstrInitialDir = startDir.c_str ();
		dlg.m_ofn.lpstrTitle = L"Load lathe parameters - pick the edited, SAVED sheet";

		if (dlg.DoModal () != IDOK)
			return 0;

		std::vector<std::filesystem::path> picked;
		for (POSITION pos = dlg.GetStartPosition (); pos != nullptr;)
			picked.push_back (std::filesystem::path (dlg.GetNextPathName (pos).GetString ()));

		// ---- Plan every file BEFORE writing anything.
		std::vector<FilePlan> plans;
		std::wstring problems;
		int changes = 0, refusals = 0, rowsChanged = 0;

		for (const std::filesystem::path &file : picked)
			{
			// The workbook as saved, or a CSV. Either way: rows of text, the
			// column names first, and each row's number as the person sees it.
			std::vector<Csv::Row> csv;
			std::vector<size_t> rowNo;
			std::wstring ext = file.extension ().wstring ();
			for (wchar_t &c : ext)
				c = static_cast<wchar_t> (towlower (c));
			if (ext == L".xlsx" || ext == L".xlsm")
				{
				std::wstring why;
				if (!Xlsx::ReadSheet (file, csv, rowNo, why))
					{
					problems += L"\r\n  " + file.filename ().wstring () + L": " + why;
					continue;
					}
				}
			else
				{
				std::wstring text;
				if (!ReadFile (file, text))
					{
					problems += L"\r\n  " + file.filename ().wstring ()
								+ L": could not be opened (is it still open in Excel?)";
					continue;
					}
				csv = Csv::Parse (text);
				for (size_t k = 0; k < csv.size (); ++k)
					rowNo.push_back (k + 1);
				}

			int typeCol = -1;
			if (!csv.empty ())
				for (size_t i = 0; i < csv[0].size (); ++i)
					if (csv[0][i] == L"type")
						typeCol = static_cast<int> (i);
			if (csv.size () < 2 || typeCol < 0)
				{
				problems += L"\r\n  " + file.filename ().wstring ()
							+ L": no header with a type column and data rows - is "
							  L"this the dumped sheet?";
				continue;
				}

			// ---- ONE SHEET, SEVERAL KINDS. Split the rows by their type, then
			// plan each kind against its own schema - so every rule the plan
			// already enforces holds unchanged. `orig` maps a row's place in the
			// split back to its line in the sheet, because a refusal that names
			// the wrong line sends somebody hunting in the wrong row.
			struct Split
				{
				const Lathe::Table *table = nullptr;
				std::vector<Csv::Row> rows;		// [0] is the header
				std::vector<size_t> orig;		// index in the sheet, aligned with rows
				};
			std::vector<Split> splits;

			for (size_t k = 1; k < csv.size (); ++k)
				{
				const std::wstring type = (static_cast<size_t> (typeCol) < csv[k].size ())
											  ? csv[k][static_cast<size_t> (typeCol)]
											  : std::wstring ();
				const Lathe::Table *t = TableForType (type);
				if (t == nullptr)
					{
					problems += L"\r\n  " + file.filename ().wstring () + L" row "
								+ std::to_wstring (rowNo[k]) + L": type \""
								+ type + L"\" is not a kind this tool loads - "
								L"the row is skipped";
					continue;
					}

				Split *s = nullptr;
				for (Split &c : splits)
					if (c.table == t)
						s = &c;
				if (s == nullptr)
					{
					splits.push_back (Split ());
					s = &splits.back ();
					s->table = t;
					s->rows.push_back (csv[0]);
					s->orig.push_back (0);
					}
				s->rows.push_back (csv[k]);
				s->orig.push_back (k);
				}

			// Header cells this sheet does not know at all, once per file.
			std::vector<std::wstring> unknown;
			for (const std::wstring &h : csv[0])
				{
				const std::wstring name = h;
				if (name.empty () || name == L"op_idn" || name == L"type"
					|| Plan::IsInfoColumn (name))
					continue;
				const auto &all = Lathe::SheetColumns ();
				if (std::find (all.begin (), all.end (), name) == all.end ())
					unknown.push_back (name);
				}

			bool firstOfFile = true;
			for (const Split &s : splits)
				{
				FilePlan fp;
				fp.file = file;
				fp.table = s.table;
				fp.result = Plan::Make (SheetSchema (*s.table), Snapshot (*s.table), s.rows);
				for (size_t k = 1; k < csv.size (); ++k)
					{
					long long id = 0;
					if (!csv[k].empty () && Csv::ParseLong (csv[k][0], id))
						fp.opRow.emplace (static_cast<long> (id), rowNo[k]);
					}

				for (Plan::Refusal &r : fp.result.refusals)
					if (r.line >= 2 && r.line - 1 < s.orig.size ())
						r.line = rowNo[s.orig[r.line - 1]];

				// Columns that belong to ANOTHER kind are not unknown - only the
				// ones no kind has are, and those are listed once.
				fp.result.unknownColumns.clear ();
				if (firstOfFile)
					fp.result.unknownColumns = unknown;
				firstOfFile = false;

				changes += static_cast<int> (fp.result.changes.size ());
				refusals += static_cast<int> (fp.result.refusals.size ());
				rowsChanged += fp.result.rowsChanged;
				plans.push_back (fp);
				}
			}

		// Everything planned goes in the log whether or not it is applied: a
		// refusal is worth being able to look at again.
		for (const FilePlan &fp : plans)
			for (const Plan::Refusal &r : fp.result.refusals)
				Util::Log (part, L"REFUSED " + fp.file.filename ().wstring ()
								 + L" row " + std::to_wstring (r.line) + L": " + r.why);

		// ---- SHOW EXACTLY WHAT WOULD HAPPEN, and let the person decide there.
		std::vector<Preview::Line> lines;
		auto add = [&lines] (Preview::Line::Kind k, const std::wstring &text,
							 const std::wstring &detail = std::wstring (),
							 const std::wstring &from = std::wstring (),
							 const std::wstring &to = std::wstring ())
			{
			Preview::Line l;
			l.kind = k;
			l.text = text;
			l.detail = detail;
			l.from = from;
			l.to = to;
			lines.push_back (l);
			};

		// Refusals first: nothing from those rows is written, and they are what
		// needs reading.
		if (refusals > 0)
			{
			add (Preview::Line::Section, L"Refused - nothing from these rows is written ("
										 + std::to_wstring (refusals) + L")");
			for (const FilePlan &fp : plans)
				for (const Plan::Refusal &r : fp.result.refusals)
					add (Preview::Line::Refused,
						 L"row " + std::to_wstring (r.line)
							 + (r.op != 0 ? L"  \u00B7  op " + std::to_wstring (r.op) : L""),
						 r.why);
			}

		// Changes grouped by operation, in the order the operations sit in the
		// sheet - which is the Operation Manager's order.
		struct Item
			{
			size_t file, row;
			long op;
			const FilePlan *fp;
			const Plan::Change *c;
			};
		std::vector<Item> items;
		for (size_t f = 0; f < plans.size (); ++f)
			for (const Plan::Change &c : plans[f].result.changes)
				{
				const auto at = plans[f].opRow.find (c.op);
				items.push_back ({ f, at == plans[f].opRow.end () ? 0 : at->second, c.op,
								   &plans[f], &c });
				}
		std::stable_sort (items.begin (), items.end (), [] (const Item &a, const Item &b)
			{ return a.file != b.file ? a.file < b.file : a.row < b.row; });

		std::map<long, std::wstring> comments;
		{
		Cnc::Tool::TpPartOpList &ops = TpMainOpMgr.GetMainOpList ();
		for (INT_PTR i = 0; i < ops.GetSize (); ++i)
			if (const operation *o = ops.GetAt (i))
				comments[o->op_idn] = std::wstring (o->comment, wcsnlen (o->comment, COMMENT_SIZE));
		}

		if (!items.empty ())
			{
			add (Preview::Line::Section, L"Changes - " + std::to_wstring (changes)
										 + (changes == 1 ? L" value on " : L" values on ")
										 + std::to_wstring (rowsChanged)
										 + (rowsChanged == 1 ? L" operation" : L" operations"));
			long lastOp = -1;
			size_t lastFile = static_cast<size_t> (-1);
			for (const Item &it : items)
				{
				if (it.op != lastOp || it.file != lastFile)
					{
					add (Preview::Line::Op,
						 L"op " + std::to_wstring (it.op) + L"  \u00B7  " + it.fp->table->schema.type
							 + (it.row ? L"  \u00B7  row " + std::to_wstring (it.row) : L""),
						 comments[it.op]);
					lastOp = it.op;
					lastFile = it.file;
					}
				add (Preview::Line::Change, it.c->name, std::wstring (), it.c->from, it.c->to);
				}
			}

		// What was set aside, last.
		if (!problems.empty ())
			{
			add (Preview::Line::Section, L"Skipped");
			size_t from = 0;
			const std::wstring sep = L"\r\n  ";
			while (from < problems.size ())
				{
				size_t at = problems.find (sep, from);
				if (at == from)
					{
					from += sep.size ();
					continue;
					}
				if (at == std::wstring::npos)
					at = problems.size ();
				add (Preview::Line::Note, problems.substr (from, at - from));
				from = at;
				}
			}
		for (const FilePlan &fp : plans)
			if (!fp.result.unknownColumns.empty ())
				{
				std::wstring cols;
				for (const std::wstring &c : fp.result.unknownColumns)
					cols += (cols.empty () ? L"" : L", ") + c;
				add (Preview::Line::Section, L"Ignored columns - not part of this tool");
				add (Preview::Line::Note, fp.file.filename ().wstring (), cols);
				}

		std::wstring files;
		for (const std::filesystem::path &f : picked)
			files += (files.empty () ? L"" : L", ") + f.filename ().wstring ();
		int unchanged = 0;
		for (const FilePlan &fp : plans)
			unchanged += fp.result.rowsUnchanged;

		const std::wstring summary =
			std::to_wstring (changes) + (changes == 1 ? L" change" : L" changes")
			+ L"  \u00B7  " + std::to_wstring (rowsChanged) + L" operation(s) changed"
			+ L"  \u00B7  " + std::to_wstring (unchanged) + L" unchanged"
			+ (refusals ? L"  \u00B7  " + std::to_wstring (refusals) + L" refused" : L"");

		if (!Preview::Show (L"Load " + files, summary, lines, changes))
			return 0;

		// ---- Apply, one database round trip per operation.
		Cnc::Tool::TpPartOpList &opList = TpMainOpMgr.GetMainOpList ();
		int wrote = 0, failed = 0;

		for (const FilePlan &fp : plans)
			{
			std::map<long, std::vector<const Plan::Change *>> byOp;
			for (const Plan::Change &c : fp.result.changes)
				byOp[c.op].push_back (&c);

			for (const auto &kv : byOp)
				{
				// THE OLD VALUES FIRST. If anything after this goes wrong, or
				// the edit turns out to be wrong, this is the record.
				for (const Plan::Change *c : kv.second)
					Util::Log (part, L"op " + std::to_wstring (kv.first) + L" "
									 + fp.table->schema.type + L"  " + c->name + L"  "
									 + OneLine (c->from) + L" -> " + OneLine (c->to));

				ent opEnt;
				if (!opList.DatabaseRetrieve (kv.first, opEnt))
					{
					++failed;
					Util::Log (part, L"op " + std::to_wstring (kv.first)
									 + L" FAILED - could not read it from the database");
					continue;
					}

				void *prm = Lathe::PrmFor (opEnt.u.op, fp.table->opcode);
				if (prm == nullptr)
					{
					++failed;
					continue;
					}

				bool ok = true;
				for (const Plan::Change *c : kv.second)
					if (!Lathe::Write (fp.table->bindings[c->col], &opEnt.u.op, prm, c->to))
						{
						ok = false;
						Util::Log (part, L"op " + std::to_wstring (kv.first)
										 + L" FAILED writing " + c->name);
						}

				if (!ok)
					{
					// Nothing is stored: opEnt is a copy, and it is discarded.
					++failed;
					continue;
					}

				// MARKED DIRTY. The post reads the NCI, and the NCI is written
				// when the operation regenerates - so an edit that is not
				// marked would sit in the operation and reach no program.
				opEnt.u.op.db.nci_flag = true;

				if (!opList.UpdateListAndDB (opEnt, true))
					{
					++failed;
					Util::Log (part, L"op " + std::to_wstring (kv.first)
									 + L" FAILED - the database rejected the update");
					continue;
					}
				++wrote;
				}
			}

		Util::Log (part, L"load: " + std::to_wstring (wrote) + L" operation(s) written, "
						 + std::to_wstring (failed) + L" failed");

		std::wstring done = L"Wrote " + std::to_wstring (wrote) + L" operation(s)";
		if (failed != 0)
			done += L", " + std::to_wstring (failed) + L" FAILED (see ParamTable.log)";
		done += L".\r\n\r\nThey are marked for regeneration - regenerate them "
				L"before posting, or the old values will still be in the "
				L"toolpath. Then save the part.";
		Util::Say (done, failed ? MB_ICONWARNING : MB_ICONINFORMATION);
		return 0;
		}
	}
