//
// LatheFields.h - which lathe toolpath parameters the tool can see, and where
// each one lives inside the SDK's parameter struct.
//
// ONE TABLE DRIVES EVERYTHING. A column is declared once - name, type, limits,
// and the pointer to its member - and the dump reads it, the comparison
// checks against it and the write applies through it. Two tables (one for
// columns, one for members) would drift, and a column bound to the wrong
// member is a wrong value written into a toolpath with no error anywhere.
//
// VALUES ARE THE RAW STORED VALUES. Nothing is converted: an angle the SDK
// stores in radians is dumped in radians, and lengths are in the part's own
// units. A round trip is only trustworthy if the dump and the load agree, and
// the easiest way to guarantee that is for neither to interpret.
//
#pragma once

#include "Plan.h"

#include <string>
#include <vector>

namespace Lathe
	{
	enum class Kind { Bool, Byte, Short, Int, Long, Double, Text,
					  Coolant,			//!< X-style coolant at one timing - see Coolant.h
					  CoolantV9,		//!< V9 coolant: the tool's one setting, by name
					  MinSec,			//!< a double of SECONDS shown and typed as m:ss
					  DoubleSize,		//!< a signed double's SIZE - the sign is kept
					  DoubleSign,		//!< a signed double's SIGN, as a word (labels)
					  LongSize,			//!< a signed long's SIZE - the sign is kept
					  LongSign,			//!< a signed long's SIGN, as a word (labels)
					  BoolWord,			//!< a bool as a word (labels: false, true)
					  ShortWord };		//!< a short holding 0, 1, 2 ... shown as wordList[n]

	/// WHERE a column's value lives. The toolpath-specific parameters (step,
	/// overlap, tool inspection) sit in the rough / finish / dynamic struct; the
	/// feeds and speeds, home position, planes and reference points sit on the
	/// OPERATION itself, shared by every kind. `at` is handed the matching base.
	enum class Base { Prm, Op };

	/// A column's home: which kind of value, and a function that finds it
	/// inside its base given a pointer to that base.
	struct Binding
		{
		Kind kind = Kind::Double;
		Base base = Base::Prm;
		void *(*at) (void *base) = nullptr;
		size_t textLen = 0;			//!< buffer size in characters, Text only
		int arg = 0;				//!< Coolant: the timing (Coolant::When)

		/// The words for a sign or a bool: [0] negative / false, [1] positive / true.
		const wchar_t *words[2] = { nullptr, nullptr };

		/// ShortWord: the word for each value, from 0.
		std::vector<std::wstring> wordList;
		};

	/// One kind of lathe operation.
	struct Table
		{
		long opcode = 0;			//!< the SDK's TP_ code
		long opcode2 = 0;			//!< a second code that means the same kind (manual entry has two)
		Plan::Schema schema;		//!< names, types and limits, in file order
		std::vector<Binding> bindings;	//!< aligned with schema.cols

		/// Where the groups begin inside schema.cols: [0, opLevelStart) is
		/// specific to this kind of toolpath, [opLevelStart, inspectStart) is the
		/// operation-level set (feeds, home, reference points, planes) and
		/// [inspectStart, end) is tool inspection. The last two are identical in
		/// every table; the sheet takes them once.
		size_t opLevelStart = 0;
		size_t inspectStart = 0;
		};

	/// EVERY column of the one sheet, in order: the toolpath-specific columns of
	/// each kind (rough, then finish, then dynamic, a name shared between kinds
	/// appearing once), then the operation-level set, then tool inspection. The
	/// identity columns (op_idn, type, tool, description) come before all of it.
	const std::vector<std::wstring> &SheetColumns ();

	/// The index of a column in this table's schema, or -1 when this kind of
	/// operation has no such column - a blank cell on the sheet.
	int IndexOf (const Table &t, const std::wstring &name);

	/// The rough, finish and dynamic rough tables.
	const std::vector<Table> &AllTables ();

	/// The table for an operation's opcode, or nullptr for any other kind.
	const Table *TableFor (long opcode);

	/// A column's value as CSV text. `op` is the operation, `prm` its toolpath
	/// parameter struct; the binding says which one it reads from.
	std::wstring Read (const Binding &b, void *op, void *prm);

	/// Write CSV text into the operation or its parameter struct. The text must
	/// already have passed Plan::Valid - this refuses only what it cannot represent.
	bool Write (const Binding &b, void *op, void *prm, const std::wstring &text);
	}

struct operation;

namespace Lathe
	{
	/// The parameter struct inside an operation for one of the three opcodes,
	/// or nullptr.
	void *PrmFor (operation &op, long opcode);
	}
