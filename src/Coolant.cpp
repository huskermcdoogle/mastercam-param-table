#include "stdafx.h"
#include "MastercamSdk.h"
#include "Coolant.h"
#include "Util.h"

#include <algorithm>
#include <cwctype>
#include <map>
#include <memory>
#include <set>

namespace
	{
	/// A machine's coolant labels as the part stores them, ten of each.
	struct Labels
		{
		std::vector<std::wstring> base, on, off;
		bool any = false;			//!< the machine names at least one coolant
		bool v9 = false;			//!< the machine uses V9 coolant (one setting, no timing)
		};

	std::map<long, Labels> gCache;	// by machine entity id, cleared by Reset

	std::wstring Trim (const std::wstring &s)
		{
		const size_t a = s.find_first_not_of (L" \t");
		if (a == std::wstring::npos)
			return std::wstring ();
		const size_t b = s.find_last_not_of (L" \t");
		return s.substr (a, b - a + 1);
		}

	std::wstring Lower (std::wstring s)
		{
		for (wchar_t &c : s)
			c = static_cast<wchar_t> (std::towlower (c));
		return s;
		}

	/// The machine definition entity an operation's group belongs to: walk up
	/// from the toolpath group until a group names one.
	long MachineOf (long grpIdn)
		{
		TpGrpList &groups = TpMainGrpMgr.GetMainGrpList ();
		long id = grpIdn;
		for (int depth = 0; depth < 32 && id > 0; ++depth)
			{
			const INT_PTR i = groups.IndexByID (id);
			if (i < 0)
				break;
			const op_group *g = groups.GetAt (i);
			if (g == nullptr)
				break;
			if (g->ogi.pg1.machine_ent_idn > 0)
				return g->ogi.pg1.machine_ent_idn;
			id = g->parent_grp_idn;
			}
		return 0;
		}

	/// Every machine definition any group in the part names.
	std::set<long> AllMachines ()
		{
		std::set<long> out;
		TpGrpList &groups = TpMainGrpMgr.GetMainGrpList ();
		for (INT_PTR i = 0; i < groups.GetSize (); ++i)
			{
			const op_group *g = groups.GetAt (i);
			if (g != nullptr && g->grp_idn > 0 && g->ogi.pg1.machine_ent_idn > 0)
				out.insert (g->ogi.pg1.machine_ent_idn);
			}
		return out;
		}

	/// A machine's labels, read once per run and logged once - the log line is
	/// how anyone checks what the part actually holds.
	const Labels &LabelsOf (long machineIdn)
		{
		const auto hit = gCache.find (machineIdn);
		if (hit != gCache.end ())
			return hit->second;

		Labels &l = gCache[machineIdn];
		l.base.assign (MAX_COOLANT_CMDS, std::wstring ());
		l.on.assign (MAX_COOLANT_CMDS, std::wstring ());
		l.off.assign (MAX_COOLANT_CMDS, std::wstring ());
		const std::filesystem::path part = Util::PartFile ();
		if (machineIdn <= 0)
			{
			Util::Log (part, L"coolant labels: no machine definition found");
			return l;
			}

		// An ent is large; keep it off the stack.
		auto e = std::make_unique<ent> ();
		bool ok = false;
		GetEntityByID (machineIdn, *e, &ok);
		if (!ok || e->id != ASSOC_ID || e->assoc_id != CNC_MACHINE_ID)
			{
			Util::Log (part, L"coolant labels: entity " + std::to_wstring (machineIdn)
							 + L" is not a machine definition");
			return l;
			}

		const coolant_type &c = e->u.machine.coolant;
		std::wstring line = L"coolant labels (machine " + std::to_wstring (machineIdn) + L"):";
		for (int k = 0; k < MAX_COOLANT_CMDS; ++k)
			{
			l.base[k] = Trim (std::wstring (c.baseTxt[k], wcsnlen (c.baseTxt[k], COOLANT_CMD_SIZE)));
			l.on[k] = Trim (std::wstring (c.onTxt[k], wcsnlen (c.onTxt[k], COOLANT_STATUS_SIZE)));
			l.off[k] = Trim (std::wstring (c.offTxt[k], wcsnlen (c.offTxt[k], COOLANT_STATUS_SIZE)));
			l.any = l.any || !l.base[k].empty ();
			if (!l.base[k].empty () || !l.on[k].empty () || !l.off[k].empty ())
				line += L"  " + std::to_wstring (k + 1) + L"=[" + l.base[k] + L" | " + l.on[k]
						+ L" | " + l.off[k] + L"]";
			}
		// "Use coolant commands in post-processor (provided for backward compatibility)":
		// the machine runs V9 coolant - the tool's one coolant setting, not canned text.
		l.v9 = c.useCoolantFromPost;
		line += l.v9 ? L"  | V9 coolant" : L"  | X-style coolant";
		Util::Log (part, line);
		return l;
		}

	/// The name the SHEET uses for each of the ten coolants: the machine's label,
	/// made safe for an Excel list (no commas or quotes) and unique; "coolant n"
	/// where the machine gives none.
	std::vector<std::wstring> Names (const Labels &l)
		{
		std::vector<std::wstring> n (MAX_COOLANT_CMDS);
		for (int k = 0; k < MAX_COOLANT_CMDS; ++k)
			{
			std::wstring s;
			for (wchar_t c : l.base[k])
				if (c == L',')
					s += L';';
				else if (c != L'"')
					s += c;
			s = Trim (s);
			n[k] = s.empty () ? L"coolant " + std::to_wstring (k + 1) : s;
			}
		for (int k = 0; k < MAX_COOLANT_CMDS; ++k)
			{
			int same = 0;
			for (int j = 0; j < MAX_COOLANT_CMDS; ++j)
				same += Lower (n[j]) == Lower (n[k]);
			if (same > 1 || Lower (n[k]) == L"none")
				n[k] += L" (" + std::to_wstring (k + 1) + L")";
			}
		return n;
		}

	struct Item
		{
		int n = 0;					//!< coolant 1..10
		bool off = false;
		};

	/// "none" / "10BAR" / "10BAR + 70BAR" / "10BAR off" against one machine's names.
	bool Parse (const std::wstring &text, const std::vector<std::wstring> &names,
				std::vector<Item> &items, std::wstring &why)
		{
		items.clear ();
		const std::wstring t = Trim (text);
		if (Lower (t) == L"none")
			return true;

		size_t start = 0;
		while (start <= t.size ())
			{
			size_t plus = t.find (L'+', start);
			if (plus == std::wstring::npos)
				plus = t.size ();
			std::wstring piece = Trim (t.substr (start, plus - start));
			start = plus + 1;

			Item it;
			const std::wstring low = Lower (piece);
			if (low.size () > 4 && low.compare (low.size () - 4, 4, L" off") == 0)
				{
				it.off = true;
				piece = Trim (piece.substr (0, piece.size () - 4));
				}
			for (size_t k = 0; k < names.size () && it.n == 0; ++k)
				if (Lower (names[k]) == Lower (piece))
					it.n = static_cast<int> (k) + 1;
			if (it.n == 0)
				{
				why = L"\"" + piece + L"\" is not a coolant on this machine (choose from:";
				for (const std::wstring &nm : names)
					why += L" " + nm + L",";
				why.back () = L')';
				return false;
				}
			items.push_back (it);
			}
		return true;
		}

	short Code (int when, const Item &it)
		{
		const int k = it.off ? 2 * it.n : 2 * it.n - 1;
		return static_cast<short> (-(when * 1000 + 100 + k));
		}
	}

namespace Coolant
	{
	void Reset ()
		{
		gCache.clear ();
		}

	std::vector<std::wstring> Choices (const operation &op)
		{
		const Labels &l = LabelsOf (MachineOf (op.cmn.grp_idn));
		const std::vector<std::wstring> names = Names (l);
		std::vector<std::wstring> out = { L"none" };
		for (int k = 0; k < MAX_COOLANT_CMDS; ++k)
			if (!l.any || !l.base[k].empty ())
				out.push_back (names[k]);
		return out;
		}

	std::wstring Describe (const operation &op, int when)
		{
		const std::vector<std::wstring> names = Names (LabelsOf (MachineOf (op.cmn.grp_idn)));
		std::wstring s;
		for (short code : op.cantxt.cantxt)
			{
			if (code >= 0)
				continue;
			const int c = -code;
			if (c / 1000 != when)
				continue;
			const int k = c % 1000 - 100;
			std::wstring item;
			if (k >= 1 && k <= 2 * MAX_COOLANT_CMDS)
				item = names[static_cast<size_t> ((k + 1) / 2 - 1)] + (k % 2 ? L"" : L" off");
			else
				item = L"code " + std::to_wstring (code);	// not ours to name; not editable
			s += (s.empty () ? L"" : L" + ") + item;
			}
		return s.empty () ? L"none" : s;
		}

	bool Apply (operation &op, int when, const std::wstring &text, std::wstring &why)
		{
		const Labels &l = LabelsOf (MachineOf (op.cmn.grp_idn));
		const std::vector<std::wstring> names = Names (l);
		std::vector<Item> items;
		if (!Parse (text, names, items, why))
			return false;
		// A V9 machine ignores X-style entries: "none" (clearing leftovers) is all it takes.
		if (l.v9 && !items.empty ())
			{
			why = L"this machine uses V9 coolant - set the coolant column instead";
			return false;
			}

		// Everything else stays: other timings' coolant and all canned text.
		std::vector<short> keep;
		for (short code : op.cantxt.cantxt)
			{
			if (code == 0)
				continue;
			if (code < 0 && (-code) / 1000 == when)
				continue;
			keep.push_back (code);
			}
		for (const Item &it : items)
			keep.push_back (Code (when, it));
		if (keep.size () > MAX_CANTXT_SIZE)
			{
			why = L"the operation holds at most " + std::to_wstring (MAX_CANTXT_SIZE)
				  + L" canned text and coolant entries";
			return false;
			}

		for (size_t i = 0; i < MAX_CANTXT_SIZE; ++i)
			op.cantxt.cantxt[i] = i < keep.size () ? keep[i] : 0;
		if (!items.empty ())
			op.cantxt.on = true;
		return true;
		}

	bool Check (const std::wstring &text, std::wstring &why)
		{
		std::set<long> machines = AllMachines ();
		if (machines.empty ())
			machines.insert (0);
		std::wstring first;
		for (long m : machines)
			{
			std::vector<Item> items;
			std::wstring w;
			if (Parse (text, Names (LabelsOf (m)), items, w))
				return true;
			if (first.empty ())
				first = w;
			}
		why = first;
		return false;
		}
	
	// ---------------------------------------------------------------- V9 coolant

	namespace
		{
		/// The V9 setting's values, as Mastercam's coolant dropdown names them.
		const struct { short bit; const wchar_t *name; } kV9[] = {
			{ 0x08, L"Off" }, { 0x10, L"Flood" }, { 0x20, L"Mist" }, { 0x40, L"Thru-tool" } };
		}

	bool IsV9 (const operation &op)
		{
		return LabelsOf (MachineOf (op.cmn.grp_idn)).v9;
		}

	bool HasXEntries (const operation &op)
		{
		for (short code : op.cantxt.cantxt)
			if (code < 0 && (-code) / 1000 >= Before && (-code) / 1000 <= After)
				return true;
		return false;
		}

	std::vector<std::wstring> ChoicesV9 ()
		{
		std::vector<std::wstring> out;
		for (const auto &v : kV9)
			out.push_back (v.name);
		return out;
		}

	std::wstring DescribeV9 (short c)
		{
		if (c == 0)
			return L"Off";
		std::wstring s;
		for (const auto &v : kV9)
			if (c & v.bit)
				s += (s.empty () ? L"" : L" + ") + std::wstring (v.name);
		if (c & ~0x78)
			s += (s.empty () ? L"" : L" + ") + std::wstring (L"code ") + std::to_wstring (c);
		return s;
		}

	bool CheckV9 (const std::wstring &text, std::wstring &why)
		{
		for (const auto &v : kV9)
			if (Lower (Trim (text)) == Lower (v.name))
				return true;
		why = L"\"" + text + L"\" is not a V9 coolant (choose from: Off, Flood, Mist, Thru-tool)";
		return false;
		}

	bool ApplyV9 (operation &op, const std::wstring &text, std::wstring &why)
		{
		if (!IsV9 (op))
			{
			why = L"this machine uses X-style coolant - set coolant_before / with / after instead";
			return false;
			}
		for (const auto &v : kV9)
			if (Lower (Trim (text)) == Lower (v.name))
				{
				op.tl.coolant = v.bit;
				return true;
				}
		return CheckV9 (text, why);
		}
	}
