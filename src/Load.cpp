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
#include "Impact.h"
#include "PartConfig.h"
#include "Undo.h"
#include "History.h"
#include "Regen.h"

#include <algorithm>
#include <fstream>
#include <iterator>
#include <map>
#include <set>
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

	/// Every operation's comment, by op_idn - the preview names operations the
	/// way the Operation Manager does.
	std::map<long, std::wstring> Comments ()
		{
		std::map<long, std::wstring> comments;
		Cnc::Tool::TpPartOpList &ops = TpMainOpMgr.GetMainOpList ();
		for (INT_PTR i = 0; i < ops.GetSize (); ++i)
			if (const operation *o = ops.GetAt (i))
				comments[o->op_idn] = std::wstring (o->comment, wcsnlen (o->comment, COMMENT_SIZE));
		return comments;
		}

	/// The values to write into one operation.
	struct Batch
		{
		const Lathe::Table *table = nullptr;
		long op = 0;
		std::vector<const Plan::Change *> changes;	//!< to = the text to write
		};

	/// Write each batch, one database round trip per operation, every old
	/// value logged first (prefixed by `prefix`, "" for a load). Counts the
	/// operations written and failed; `written` (when given) gets the ones written.
	void WriteBatches (const std::filesystem::path &part, const std::vector<Batch> &batches,
				const std::wstring &prefix, int &wrote, int &failed, std::vector<long> *written = nullptr)
		{
		Cnc::Tool::TpPartOpList &opList = TpMainOpMgr.GetMainOpList ();
		wrote = failed = 0;

		for (const Batch &b : batches)
			{
			// THE OLD VALUES FIRST. If anything after this goes wrong, or
			// the edit turns out to be wrong, this is the record.
			for (const Plan::Change *c : b.changes)
				Util::Log (part, prefix + Undo::ChangeLine (b.op, b.table->schema.type, c->name,
															c->from, c->to));

			ent opEnt;
			if (!opList.DatabaseRetrieve (b.op, opEnt))
				{
				++failed;
				Util::Log (part, prefix + L"op " + std::to_wstring (b.op)
								 + L" FAILED - could not read it from the database");
				continue;
				}

			void *prm = Lathe::PrmFor (opEnt.u.op, b.table->opcode);
			if (prm == nullptr)
				{
				++failed;
				continue;
				}

			bool ok = true;
			for (const Plan::Change *c : b.changes)
				if (!Lathe::Write (b.table->bindings[c->col], &opEnt.u.op, prm, c->to))
					{
					ok = false;
					Util::Log (part, prefix + L"op " + std::to_wstring (b.op)
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
				Util::Log (part, prefix + L"op " + std::to_wstring (b.op)
								 + L" FAILED - the database rejected the update");
				continue;
				}
			++wrote;
			if (written != nullptr && std::find (written->begin (), written->end (), b.op) == written->end ())
				written->push_back (b.op);
			}
		}

	/// A workbook loaded, as far as the part's history needs it: its sheets, for
	/// the insert cost per part before and after.
	struct Loaded
		{
		std::filesystem::path file;
		std::vector<Csv::Row> main, dumped;
		Xlsx::Grid tools;
		};

	/// A preview line.
	Preview::Line MakeLine (Preview::Line::Kind k, const std::wstring &text,
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
		return l;
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
		// Each operation's time and flips, as dumped and as edited - the sheet's
		// own estimate columns, which the plan never reads.
		std::map<long, Impact::Op> figures;
		// The part's history (<part>.pthistory), the workbooks as it needs them, and
		// the batch quantity typed in them.
		History::Book history = History::Open (part);
		std::vector<Loaded> loaded;
		std::wstring batch;

		for (const std::filesystem::path &file : picked)
			{
			// The workbook as saved, or a CSV. Either way: rows of text, the
			// column names first, and each row's number as the person sees it.
			std::vector<Csv::Row> csv;
			std::vector<size_t> rowNo;
			std::vector<Csv::Row> dumped;
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
				// The values as dumped, for what the edits moved the flips FROM.
				// A sheet without them still loads; its flips compare to `flips`.
				std::vector<size_t> dumpedRow;
				if (!Xlsx::ReadNamedSheet (file, L"Dumped", dumped, dumpedRow, why))
					dumped.clear ();
				// What was typed for the part's inserts and tools (edges, costs, edge
				// life, batch) goes into its .ptconfig, for the next dump to start from.
				{
				const std::filesystem::path cfgFile = PartConfig::PathFor (part);
				PartConfig::Config cfg = PartConfig::Load (cfgFile);
				std::wstring cfgWhy;
				if (PartConfig::Harvest (file, cfg, cfgWhy) && PartConfig::Save (cfgFile, cfg))
					Util::Log (part, L"part config: took what was typed in " + file.filename ().wstring ());
				batch = cfg.batchQty;
				}
				// What was measured on the machine (the Summary's "From the machine"
				// cells) goes into the part's history - each measurement once.
				{
				const std::vector<History::Record> measured = History::Measured (file, history.records, History::Now ());
				if (!measured.empty ())
					Util::Log (part, history.Add (measured) ? L"history: a measurement from " + file.filename ().wstring ()
															: L"history: could not write " + history.file.wstring ());
				}
				Loaded l;
				l.file = file;
				l.main = csv;
				l.dumped = dumped;
				Xlsx::ReadGrid (file, L"Tools", l.tools, why);
				loaded.push_back (l);
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

			for (const auto &kv : Impact::Read (csv, dumped))
				figures[kv.first] = kv.second;

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
		// (tests\dialog_shots.cpp puts the lines together the same way, for the
		// manual's picture of this window: a change here wants one there.)
		std::vector<Preview::Line> lines;
		auto add = [&lines] (Preview::Line::Kind k, const std::wstring &text,
							 const std::wstring &detail = std::wstring (),
							 const std::wstring &from = std::wstring (),
							 const std::wstring &to = std::wstring ()) -> Preview::Line &
			{
			lines.push_back (MakeLine (k, text, detail, from, to));
			return lines.back ();
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

		std::map<long, std::wstring> comments = Comments ();

		if (!items.empty ())
			{
			add (Preview::Line::Section, L"Changes - " + std::to_wstring (changes)
										 + (changes == 1 ? L" value on " : L" values on ")
										 + std::to_wstring (rowsChanged)
										 + (rowsChanged == 1 ? L" operation" : L" operations"));
			long lastOp = -1;
			size_t lastFile = static_cast<size_t> (-1);
			for (size_t k = 0; k < items.size (); ++k)
				{
				const Item &it = items[k];
				if (it.op != lastOp || it.file != lastFile)
					{
					Preview::Line &o = add (Preview::Line::Op,
											L"op " + std::to_wstring (it.op) + L"  \u00B7  " + it.fp->table->schema.type
												+ (it.row ? L"  \u00B7  row " + std::to_wstring (it.row) : L""),
											comments[it.op]);
					o.box = Preview::Line::Ticked;
					o.tag = it.op;
					const auto f = figures.find (it.op);
					if (f != figures.end ())
						{
						o.impact = Impact::OpText (f->second);
						o.tone = Impact::Tone (f->second.was, f->second.now);
						if (o.tone == 0)
							o.tone = Impact::Tone (f->second.flipsWas, f->second.flipsNow, true);
						}
					lastOp = it.op;
					lastFile = it.file;
					}
				Preview::Line &c = add (Preview::Line::Change, it.c->name, std::wstring (), it.c->from, it.c->to);
				c.box = Preview::Line::Ticked;
				c.link = Plan::LinkOf (it.c->name);
				c.tag = static_cast<long> (k);
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

		// ---- THE IMPACT, following the ticks: an operation counts as edited
		// while any of its changes is ticked, as dumped when none is.
		Preview::Options options;
		options.foot = L"Untick a change to leave it out. Every old value goes to ParamTable.log first - "
					   L"\"Lathe params - undo last load\" puts them back. Changed operations are marked "
					   L"for regeneration.";
		if (!figures.empty ())
			options.impact = [&figures] (std::vector<Preview::Line> &ls, int &tone) -> std::wstring
				{
				std::set<long> applied;
				bool partly = false;
				for (const Preview::Line &l : ls)
					if (l.kind == Preview::Line::Op && l.box != Preview::Line::Unticked)
						{
						applied.insert (l.tag);
						partly = partly || l.box == Preview::Line::Mixed;
						}
				const Impact::Total t = Impact::Sum (figures, applied);
				std::wstring s = Impact::TotalText (t);
				if (s.empty ())
					return s;
				tone = Impact::Tone (t.was, t.now);
				if (tone == 0)
					tone = Impact::Tone (t.flipsWas, t.flipsNow, true);
				// The estimate is per row: a row half applied cannot be split.
				if (partly)
					s += L"  (partly ticked operations count all their edits)";
				return s;
				};
		// Regenerating afterwards: a choice made here each time, unticked to start -
		// on a big part it takes twenty minutes.
		bool regenAfter = false;
		options.choice = [] (const std::vector<Preview::Line> &ls)
			{
			int n = 0;
			for (const Preview::Line &l : ls)
				if (l.kind == Preview::Line::Op && l.box != Preview::Line::Unticked)
					++n;
			return L"Regenerate the " + std::to_wstring (n) + (n == 1 ? L" changed operation" : L" changed operations")
				   + L" after loading - can take a long time on big parts";
			};
		options.chosen = &regenAfter;

		if (!Preview::Show (L"Load " + files, summary, lines, options))
			return 0;

		// ---- What was ticked. A change left out is logged, so the log still
		// tells the whole story of the sheet.
		std::set<const Plan::Change *> chosen;
		for (const Preview::Line &l : lines)
			{
			if (l.kind != Preview::Line::Change || l.tag < 0 || static_cast<size_t> (l.tag) >= items.size ())
				continue;
			const Item &it = items[static_cast<size_t> (l.tag)];
			if (l.box == Preview::Line::Ticked)
				chosen.insert (it.c);
			else
				Util::Log (part, L"left out: " + Undo::ChangeLine (it.op, it.fp->table->schema.type, it.c->name,
																 it.c->from, it.c->to));
			}

		// ---- Apply, one database round trip per operation.
		std::vector<Batch> batches;
		for (const FilePlan &fp : plans)
			{
			std::map<long, std::vector<const Plan::Change *>> byOp;
			for (const Plan::Change &c : fp.result.changes)
				if (chosen.count (&c) != 0)
					byOp[c.op].push_back (&c);
			for (const auto &kv : byOp)
				{
				Batch b;
				b.table = fp.table;
				b.op = kv.first;
				b.changes = kv.second;
				batches.push_back (b);
				}
			}

		// The line that makes this load undoable: it names the part, because one
		// log serves every part in the folder.
		Util::Log (part, Undo::BeginLine (part.filename ().wstring (), files));
		int wrote = 0, failed = 0;
		std::vector<long> written;
		WriteBatches (part, batches, std::wstring (), wrote, failed, &written);

		Util::Log (part, L"load: " + std::to_wstring (wrote) + L" operation(s) written, "
						 + std::to_wstring (failed) + L" failed");

		// ---- THE PART'S HISTORY: what this load did to the part's figures - before
		// as dumped, after as the sheet showed, for the operations written - each
		// op's estimate (what a regeneration is measured against), and the reasons
		// typed in the workbook's Change report.
		const std::set<long> applied (written.begin (), written.end ());
		const Impact::Total total = Impact::Sum (figures, applied);
		{
		History::LoadFigures f;
		f.files = files;
		f.changes = static_cast<int> (chosen.size ());
		f.ops = wrote;
		f.failed = failed;
		if (total.timed > 0)
			{
			f.cycleWas = total.was;
			f.cycleNow = total.now;
			}
		if (total.flipped > 0)
			{
			f.flipsWas = total.flipsWas;
			f.flipsNow = total.flipsNow;
			}
		// Insert cost from the workbook's own Tools page - one workbook only: two
		// would each have their own costs and their own share of the ops.
		if (loaded.size () == 1)
			{
			f.costWas = History::InsertCost (loaded[0].main, loaded[0].dumped, loaded[0].tools, {});
			f.costNow = History::InsertCost (loaded[0].main, loaded[0].dumped, loaded[0].tools, applied);
			}
		f.batch = batch;
		for (long op : written)
			{
			const auto at = figures.find (op);
			f.estimates.push_back ({ op, at != figures.end () ? at->second.now : std::nan ("") });
			}
		const std::wstring now = History::Now ();
		std::vector<History::Record> recs = { History::LoadRecord (f, now) };
		for (const Loaded &l : loaded)
			for (const History::Record &w : History::Whys (l.file, history.records, now))
				recs.push_back (w);
		if (!history.Add (recs))
			Util::Log (part, L"history: could not write " + history.file.wstring ());
		}

		// ---- REGENERATE, only when it was ticked in the preview.
		std::wstring regenSaid;
		if (regenAfter && !written.empty ())
			{
			std::map<long, double> estimates;
			for (long op : written)
				{
				const auto at = figures.find (op);
				if (at != figures.end ())
					estimates[op] = at->second.now;
				}
			Regen::Outcome o;
			{
			CWaitCursor wait;
			o = Regen::Run (part, written, estimates);
			}
			Regen::Keep (part, history, o, total.timed > 0 ? total.now : std::nan (""));
			regenSaid = Regen::Said (o);
			}

		std::wstring done = L"Wrote " + std::to_wstring (wrote) + L" operation(s)";
		if (failed != 0)
			done += L", " + std::to_wstring (failed) + L" FAILED (see ParamTable.log)";
		if (!regenSaid.empty ())
			done += L".\r\n\r\n" + regenSaid + L"\r\n\r\nThen save the part.";
		else
			done += L".\r\n\r\nThey are marked for regeneration - regenerate them "
					L"before posting, or the old values will still be in the "
					L"toolpath (\"Lathe params - regenerate last load\" does it, and "
					L"checks the estimate against Mastercam). Then save the part.";
		done += L"\r\n\r\nTo put the old values back: \"Lathe params - undo last load\".";
		Util::Say (done, failed ? MB_ICONWARNING : MB_ICONINFORMATION);
		return 0;
		}

	int UndoLast ()
		{
		Coolant::Reset ();
		const std::filesystem::path part = Util::PartFile ();
		if (part.empty ())
			{
			Util::Say (L"Save the part first - there is no open part to undo a load in.");
			return 0;
			}
		const std::wstring name = part.filename ().wstring ();

		// ---- The last load of this part, from the log beside it.
		std::wstring log;
		if (!ReadFile (part.parent_path () / L"ParamTable.log", log))
			{
			Util::Say (L"There is no ParamTable.log beside " + name + L", so there is no load to undo.");
			return 0;
			}
		const Undo::Last last = Undo::FindLast (log, name);
		if (!last.found)
			{
			Util::Say (L"ParamTable.log has no load of " + name + L" that can be undone."
					   + (last.olderLoads
							  ? std::wstring (L"\r\n\r\nIts loads were made by an older version of the tool, "
											  L"which did not note the part. Their old values are still in the "
											  L"log, to put back by hand.")
							  : std::wstring ()),
					   MB_ICONINFORMATION);
			return 0;
			}

		// ---- Each value still as the load left it goes back.
		std::vector<Undo::Kind> kinds;
		for (const Lathe::Table &t : Lathe::AllTables ())
			{
			Undo::Kind k;
			k.schema = &t.schema;
			k.now = Snapshot (t);
			kinds.push_back (k);
			}
		const Undo::Result r = Undo::Make (last, kinds);

		// The window's lines (tests\dialog_shots.cpp makes them the same way for the manual).
		std::vector<Preview::Line> lines;
		if (!r.skipped.empty ())
			{
			lines.push_back (MakeLine (Preview::Line::Section, L"Changed since the load - left as they are ("
									   + std::to_wstring (r.skipped.size ()) + L")"));
			for (const Undo::Skip &s : r.skipped)
				lines.push_back (MakeLine (Preview::Line::Refused,
										   L"op " + std::to_wstring (s.op) + L"  ·  " + s.type, s.why));
			}

		std::map<long, std::wstring> comments = Comments ();
		int ops = 0;
		{
		long lastOp = -1;
		for (const Undo::Restore &x : r.restores)
			if (x.change.op != lastOp)
				{
				++ops;
				lastOp = x.change.op;
				}
		}
		if (!r.restores.empty ())
			{
			lines.push_back (MakeLine (Preview::Line::Section, L"Restore - " + std::to_wstring (r.restores.size ())
									   + (r.restores.size () == 1 ? L" value on " : L" values on ")
									   + std::to_wstring (ops) + (ops == 1 ? L" operation" : L" operations")));
			long lastOp = -1;
			for (size_t k = 0; k < r.restores.size (); ++k)
				{
				const Undo::Restore &x = r.restores[k];
				if (x.change.op != lastOp)
					{
					Preview::Line o = MakeLine (Preview::Line::Op, L"op " + std::to_wstring (x.change.op)
												+ L"  ·  " + x.type, comments[x.change.op]);
					o.box = Preview::Line::Ticked;
					o.tag = x.change.op;
					lines.push_back (o);
					lastOp = x.change.op;
					}
				Preview::Line c = MakeLine (Preview::Line::Change, x.change.name, std::wstring (),
											x.change.from, x.change.to);
				c.box = Preview::Line::Ticked;
				c.link = Plan::LinkOf (x.change.name);
				c.tag = static_cast<long> (k);
				lines.push_back (c);
				}
			}

		if (!r.already.empty ())
			{
			lines.push_back (MakeLine (Preview::Line::Section, L"Already as before the load ("
									   + std::to_wstring (r.already.size ()) + L")"));
			for (const Undo::Entry &e : r.already)
				lines.push_back (MakeLine (Preview::Line::Note, L"op " + std::to_wstring (e.op) + L"  ·  "
										   + e.type, e.column));
			}

		const std::wstring summary =
			L"Load of " + last.stamp + (last.files.empty () ? L"" : L" from " + last.files)
			+ L"  ·  " + std::to_wstring (r.restores.size ()) + L" to restore"
			+ (r.skipped.empty () ? L"" : L"  ·  " + std::to_wstring (r.skipped.size ()) + L" changed since")
			+ (r.already.empty () ? L"" : L"  ·  " + std::to_wstring (r.already.size ()) + L" already back");

		Preview::Options options;
		options.caption = L"Parameter Table Tool - undo last load";
		options.verb = L"Restore";
		options.foot = L"Untick a value to keep it as it is now. Every value replaced goes to ParamTable.log "
					   L"first. Restored operations are marked for regeneration.";
		options.nothing = last.entries.empty () ?L"That load wrote nothing - there is nothing to restore."
											: L"Nothing to restore.";

		if (!Preview::Show (L"Undo the last load of " + name, summary, lines, options))
			return 0;

		// ---- Restore, logged the way a load is, marked as an undo.
		Util::Log (part, L"undo begin: " + name + L"  (load of " + last.stamp + L")");
		std::map<std::pair<long, std::wstring>, Batch> byOp;
		std::vector<std::pair<long, std::wstring>> order;
		for (const Preview::Line &l : lines)
			{
			if (l.kind != Preview::Line::Change || l.tag < 0 || static_cast<size_t> (l.tag) >= r.restores.size ())
				continue;
			const Undo::Restore &x = r.restores[static_cast<size_t> (l.tag)];
			if (l.box != Preview::Line::Ticked)
				{
				Util::Log (part, L"undo left out: " + Undo::ChangeLine (x.change.op, x.type, x.change.name,
																	  x.change.from, x.change.to));
				continue;
				}
			const Lathe::Table *t = TableForType (x.type);
			if (t == nullptr)
				continue;
			const auto key = std::make_pair (x.change.op, x.type);
			if (byOp.count (key) == 0)
				{
				order.push_back (key);
				byOp[key].table = t;
				byOp[key].op = x.change.op;
				}
			byOp[key].changes.push_back (&x.change);
			}
		std::vector<Batch> batches;
		for (const auto &key : order)
			batches.push_back (byOp[key]);

		int wrote = 0, failed = 0;
		WriteBatches (part, batches, L"undo ", wrote, failed);
		Util::Log (part, L"undo: " + std::to_wstring (wrote) + L" operation(s) restored, "
						 + std::to_wstring (failed) + L" failed");

		// The part's history: the load taken back - all of it (its saving then no
		// longer counts), or only part.
		{
		size_t values = 0, offered = 0;
		for (const Batch &b : batches)
			values += b.changes.size ();
		for (const Preview::Line &l : lines)
			if (l.kind == Preview::Line::Change)
				++offered;
		History::Record u;
		u.when = History::Now ();
		u.kind = L"undo";
		u.Set (L"load", last.stamp).Set (L"restored", std::to_wstring (values));
		u.Set (L"complete", r.skipped.empty () && failed == 0 && values == offered ? L"yes" : L"no");
		History::Book history = History::Open (part);
		if (!history.Add ({ u }))
			Util::Log (part, L"history: could not write " + history.file.wstring ());
		}

		std::wstring done = L"Restored " + std::to_wstring (wrote) + L" operation(s)";
		if (failed != 0)
			done += L", " + std::to_wstring (failed) + L" FAILED (see ParamTable.log)";
		done += L".\r\n\r\nThey are marked for regeneration - regenerate them "
				L"before posting. Then save the part.";
		Util::Say (done, failed ? MB_ICONWARNING : MB_ICONINFORMATION);
		return 0;
		}
	}
