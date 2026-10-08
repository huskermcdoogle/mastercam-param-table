//
// plan_test.cpp - what a reloaded CSV may change, outside Mastercam.
//
// Each block is one of the rules in Plan.h, because each rule is a silent
// wrong value if it is broken.
//
#include "../src/Plan.h"

#include <cstdio>

namespace
	{
	int gFailed = 0;

	void Check (bool ok, const char *what)
		{
		if (!ok)
			{
			std::printf ("FAIL  %s\n", what);
			++gFailed;
			}
		}

	Plan::Schema RoughSchema ()
		{
		Plan::Schema s;
		s.type = L"ROUGH";

		Plan::Col dir;   dir.name = L"direction";   dir.type = Plan::Type::Long;
		dir.readOnly = true;
		Plan::Col step;  step.name = L"step";       step.type = Plan::Type::Double;
		step.lo = 0.0;
		Plan::Col eq;    eq.name = L"equal_steps";  eq.type = Plan::Type::Bool;
		Plan::Col cuts;  cuts.name = L"fin_n_cuts"; cuts.type = Plan::Type::Long;
		cuts.lo = 0; cuts.hi = 99;
		Plan::Col cmt;   cmt.name = L"insp_comment"; cmt.type = Plan::Type::Text;
		cmt.hi = 12;			// a 12-character buffer, for the test

		s.cols = { dir, step, eq, cuts, cmt };
		return s;
		}

	std::vector<Plan::Current> Ops ()
		{
		std::vector<Plan::Current> v;
		Plan::Current a;
		a.op = 10; a.type = L"ROUGH";
		a.values = { L"0", L"0.1", L"1", L"2", L"check tool" };
		Plan::Current b;
		b.op = 11; b.type = L"ROUGH";
		b.values = { L"1", L"0.05", L"0", L"0", L"" };
		Plan::Current f;
		f.op = 20; f.type = L"FINISH";
		f.values = { L"0", L"0.01", L"0", L"0", L"" };
		v.push_back (a);
		v.push_back (b);
		v.push_back (f);
		return v;
		}

	std::vector<Csv::Row> Sheet (std::initializer_list<Csv::Row> rows)
		{
		return std::vector<Csv::Row> (rows);
		}

	const Csv::Row kHead = { L"op_idn", L"type", L"direction", L"step",
							 L"equal_steps", L"fin_n_cuts", L"insp_comment" };
	}

int main ()
	{
	const Plan::Schema schema = RoughSchema ();
	const std::vector<Plan::Current> ops = Ops ();

	// ---- ONLY A DIFFERING CELL IS A CHANGE.
	{
	// The sheet exactly as dumped, with Excel's habits: 0.10 for 0.1, TRUE
	// for 1, 2.0 for 2.
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.10", L"TRUE", L"2.0", L"check tool" },
		{ L"11", L"ROUGH", L"1", L"0.05", L"FALSE", L"0", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty (), "an unedited sheet changes nothing");
	Check (r.refusals.empty (), "an unedited sheet is not refused");
	Check (r.rowsUnchanged == 2, "both rows are counted as unchanged");
	}

	// ---- A real edit is one change with from and to.
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.15", L"1", L"2", L"check tool" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.size () == 1, "one edited cell is one change");
	Check (!r.changes.empty () && r.changes[0].name == L"step"
		   && r.changes[0].from == L"0.1" && r.changes[0].to == L"0.15"
		   && r.changes[0].op == 10, "the change says what, from and to");
	Check (r.rowsChanged == 1, "the row counts as changed");
	}

	// ---- A CELL LEFT EMPTY MEANS LEAVE IT ALONE, NEVER ZERO.
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"", L"", L"", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty (), "deleted cells are not zeroed");
	Check (r.emptyCells == 4, "and the deleted cells are counted, not hidden");
	}

	// ---- COLUMN ORDER AND MISSING COLUMNS.
	{
	const auto csv = Sheet ({
		{ L"step", L"type", L"op_idn" },
		{ L"0.2", L"ROUGH", L"10" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.size () == 1 && r.changes[0].to == L"0.2",
		   "columns may be reordered and some left out");
	}

	// ---- A ROW IS CHECKED AGAINST ITS TYPE.
	{
	const auto csv = Sheet ({ kHead,
		{ L"20", L"ROUGH", L"0", L"0.5", L"0", L"0", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty (), "a rough row cannot change a finish operation");
	Check (r.refusals.size () == 1, "and it is refused, not skipped silently");
	}
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"FINISH", L"0", L"0.5", L"0", L"0", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "a row labelled with the wrong type is refused");
	}
	{
	const auto csv = Sheet ({ kHead,
		{ L"99", L"ROUGH", L"0", L"0.5", L"0", L"0", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "an operation that is not in the part is refused");
	}

	// ---- A BAD CELL REFUSES THE WHOLE ROW.
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.9", L"1", L"abc", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty (),
		   "a good cell beside a bad one is NOT applied - half a row is worse");
	Check (r.refusals.size () == 1, "the bad cell is named");
	}
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"-1", L"1", L"2", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "below the minimum is refused");
	}
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.1", L"1", L"100", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "above the maximum is refused");
	}
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"#REF!", L"1", L"2", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "an Excel error value is refused");
	}

	// ---- One bad row does not take its neighbours with it.
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.9", L"1", L"abc", L"" },
		{ L"11", L"ROUGH", L"1", L"0.08", L"0", L"0", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.size () == 1 && r.changes[0].op == 11,
		   "a good row still applies beside a refused one");
	Check (r.refusals.size () == 1 && r.refusals[0].line == 2,
		   "and the refusal names the right line");
	}

	// ---- READ-ONLY COLUMNS.
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"1", L"0.1", L"1", L"2", L"check tool" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty (), "a read-only column is never written");
	Check (r.refusals.size () == 1, "editing one is reported");
	}

	// ---- TEXT THAT WOULD BE CUT OFF is refused, not truncated.
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.1", L"1", L"2", L"this is far too long" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "a comment longer than the operation's buffer is refused");
	}
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.1", L"1", L"2", L"fits fine" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.size () == 1 && r.changes[0].name == L"insp_comment",
		   "a comment that fits is a change");
	}

	// ---- ONE SHEET, EVERY KIND: a value in another kind's column is refused.
	{
	Plan::Schema s = schema;
	s.foreign = { L"n_cuts" };			// a finish-only column on a rough sheet

	const Csv::Row head = { L"op_idn", L"type", L"step", L"n_cuts" };

	// Blank, as a dump leaves it: not a problem, not a change.
	const auto blank = Sheet ({ head, { L"10", L"ROUGH", L"0.1", L"" } });
	const Plan::Result a = Plan::Make (s, ops, blank);
	Check (a.changes.empty () && a.refusals.empty (),
		   "a blank cell in another kind's column is fine");

	// Typed: it would change nothing, and the person must be told.
	const auto typed = Sheet ({ head, { L"10", L"ROUGH", L"0.2", L"3" } });
	const Plan::Result b = Plan::Make (s, ops, typed);
	Check (b.changes.empty (),
		   "a value in another kind's column refuses the row, edits and all");
	Check (b.refusals.size () == 1, "and says so");

	// A finish row may use its own column - the same header, a different type.
	const auto other = Sheet ({ head, { L"10", L"ROUGH", L"0.1", L"  " } });
	const Plan::Result c = Plan::Make (s, ops, other);
	Check (c.refusals.empty (), "whitespace alone is blank");
	}

	// ---- The identity columns are not "unknown" columns.
	{
	const auto csv = Sheet ({
		{ L"op_idn", L"type", L"tool", L"comment", L"step" },
		{ L"10", L"ROUGH", L"5", L"Rough OD", L"0.1" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.unknownColumns.empty (), "tool and comment are identity columns, not unknown");
	}

	// ---- Duplicates.
	{
	const auto csv = Sheet ({ kHead,
		{ L"10", L"ROUGH", L"0", L"0.2", L"1", L"2", L"check tool" },
		{ L"10", L"ROUGH", L"0", L"0.3", L"1", L"2", L"check tool" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.size () == 1 && r.changes[0].to == L"0.2",
		   "the first row for an operation wins");
	Check (r.refusals.size () == 1, "the repeat is refused, not merged");
	}

	// ---- Whole-file problems.
	{
	const Plan::Result r = Plan::Make (schema, ops, {});
	Check (r.changes.empty () && r.refusals.size () == 1, "an empty file");
	}
	{
	const auto csv = Sheet ({ { L"step", L"fin_n_cuts" }, { L"0.1", L"2" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "no op_idn column means nothing can be matched");
	}
	{
	const auto csv = Sheet ({ { L"op_idn", L"type", L"step", L"step" },
							  { L"10", L"ROUGH", L"0.1", L"0.2" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1,
		   "a repeated column name is refused rather than guessed at");
	}

	// ---- Columns the schema does not know are listed, not obeyed.
	{
	const auto csv = Sheet ({
		{ L"op_idn", L"type", L"step", L"my_notes" },
		{ L"10", L"ROUGH", L"0.1", L"hello" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.unknownColumns.size () == 1 && r.unknownColumns[0] == L"my_notes",
		   "an extra column is reported and ignored");
	Check (r.changes.empty (), "and does not change anything");
	}

	// ---- A row that names a non-number as its operation.
	{
	const auto csv = Sheet ({ kHead,
		{ L"ten", L"ROUGH", L"0", L"0.2", L"1", L"2", L"" } });
	const Plan::Result r = Plan::Make (schema, ops, csv);
	Check (r.changes.empty () && r.refusals.size () == 1 && r.refusals[0].op == 0,
		   "an unreadable op_idn is refused with no operation attached");
	}

	// ---- A column of fixed choices (feed_mode, spindle_dir ...).
	{
	Plan::Col c;
	c.name = L"feed_mode";
	c.type = Plan::Type::Text;
	c.choices = { L"per rev", L"per min" };
	std::wstring why;
	Check (Plan::Valid (c, L"per rev", why), "a choice is valid");
	Check (Plan::Valid (c, L" Per Min ", why), "a choice in another case, padded, is valid");
	Check (!Plan::Valid (c, L"ipr", why) && !why.empty (), "anything else is refused, with a reason");
	Check (Plan::SameValue (c, L"per rev", L"PER REV"), "a choice is the same in any case");
	Check (!Plan::SameValue (c, L"per rev", L"per min"), "two choices differ");
	}

	if (gFailed != 0)
		{
		std::printf ("plan_test: %d check(s) FAILED\n", gFailed);
		return 1;
		}
	std::printf ("plan_test: all checks passed\n");
	return 0;
	}
