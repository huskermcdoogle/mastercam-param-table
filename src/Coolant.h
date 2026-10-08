//
// Coolant.h - X-style coolant as three sheet columns: which coolants come on
// BEFORE, WITH and AFTER the move, by the machine's own coolant names.
//
// WHERE IT LIVES. X-style coolant is stored in the operation's canned-text
// array (op.cantxt), in the same 20 slots as ordinary canned text (M00, M01,
// tailstock ...). A coolant entry is a NEGATIVE code, confirmed on a test part:
//
//     -(when * 1000 + 100 + k)
//
//   when = 3 before, 4 with, 5 after the move - the post's cant_pos;
//   k    = the command's place in the post's coolant list, which the post
//          receives as cantext 49 + k: odd k turns coolant (k + 1) / 2 ON,
//          even k turns coolant k / 2 OFF.
//
// So coolant 1 on with the move is -4101, and coolant 7 on after it is -5113.
// A write replaces only the coolant entries of ONE timing and leaves every
// other entry - other timings, ordinary canned text - exactly as it was.
//
// V9 COOLANT. A machine definition can instead use V9 coolant ("use coolant
// commands in post-processor", kept for backward compatibility): then the
// operation has ONE coolant setting - Off, Flood, Mist or Thru-tool - in the
// tool info's bit field (0x08 / 0x10 / 0x20 / 0x40), with no before / with /
// after, and the X-style entries above are ignored.
//
// NAMES come from the PART: the operation's machine group carries a machine
// definition entity, whose coolant labels are what the Coolant dialog shows.
//
#pragma once

#include <string>
#include <vector>

struct operation;

namespace Coolant
	{
	enum When { Before = 3, With = 4, After = 5 };

	/// Forget cached machine labels - call at the start of every dump and load.
	void Reset ();

	/// What the dropdown offers for this operation's machine: "none" first,
	/// then every coolant the machine names (all ten when it names none).
	std::vector<std::wstring> Choices (const operation &op);

	/// The coolants this operation turns on (or off) at one timing, as the
	/// sheet shows them: "10BAR", "10BAR + 70BAR", "10BAR off", or "none".
	std::wstring Describe (const operation &op, int when);

	/// Replace this timing's coolant entries with the ones `text` names.
	/// False, with a reason, when a name is not this machine's or the array
	/// would overflow - and then the operation is not touched.
	bool Apply (operation &op, int when, const std::wstring &text, std::wstring &why);

	/// Plan-time check, before anything is written: every name in `text` is a
	/// coolant of SOME machine in this part. Apply checks the exact machine.
	bool Check (const std::wstring &text, std::wstring &why);

	/// The operation's machine uses V9 coolant.
	bool IsV9 (const operation &op);

	/// The operation carries X-style coolant entries (on a V9 machine: leftovers).
	bool HasXEntries (const operation &op);

	/// V9: Off, Flood, Mist, Thru-tool.
	std::vector<std::wstring> ChoicesV9 ();

	/// V9: the bit field in words ("Flood"; "Off" for 0).
	std::wstring DescribeV9 (short coolant);

	/// V9: the text is one of the four names.
	bool CheckV9 (const std::wstring &text, std::wstring &why);

	/// V9: set the operation's one coolant setting. False on an X-style machine.
	bool ApplyV9 (operation &op, const std::wstring &text, std::wstring &why);
	}
