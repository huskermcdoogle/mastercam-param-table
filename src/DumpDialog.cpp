#include "stdafx.h"
#include "MastercamSdk.h"
#include "DumpDialog.h"
#include "FileRules.h"
#include "Pick.h"

#include <algorithm>
#include <ctime>
#include <filesystem>
#include <map>
#include <set>

namespace
	{
	const COLORREF kInk   = RGB (0x1F, 0x29, 0x37);
	const COLORREF kMuted = RGB (0x5B, 0x65, 0x73);
	const COLORREF kGood  = RGB (0x06, 0x76, 0x47);
	const COLORREF kFound = RGB (0xFF, 0xEB, 0x9C);		// an operation the search found
	const COLORREF kFaded = RGB (0xA0, 0xA6, 0xAF);		// one it did not

	enum
		{
		IdTitle = 2001, IdSummary, IdOpsHead, IdTree, IdAll, IdNone, IdMgr, IdKind, IdKindOn, IdKindOff,
		IdSkipped, IdFolderHead, IdFolder, IdBrowse, IdBeside, IdNameHead, IdPattern, IdPreview,
		IdTokens, IdOpen, IdPictures, IdMacros, IdStockSim, IdFindHead, IdFind, IdFindOn, IdFindOff, IdFindCount,
		IdSavedHead, IdSaved, IdSave, IdDelete, IdSavedNote
		};

	class Dlg : public CDialog
		{
		public:
			Dlg (const std::wstring &part, std::vector<DumpDialog::Op> &ops, const std::wstring &skipped,
				 Settings::Dump &settings, CWnd *parent)
				: m_part (part), m_ops (ops), m_skipped (skipped), m_settings (settings), m_parent (parent)
				{
				}

			INT_PTR Run ()
				{
				m_tpl.assign (64, 0);
				DLGTEMPLATE *t = reinterpret_cast<DLGTEMPLATE *> (m_tpl.data ());
				t->style = WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME | DS_CENTER;
				t->cx = 300;
				t->cy = 300;
				InitModalIndirect (t, m_parent);
				return DoModal ();
				}

		protected:
			BOOL OnInitDialog () override
				{
				CDialog::OnInitDialog ();
				SetWindowText (L"Parameter Table Tool - dump");

				NONCLIENTMETRICS ncm = { sizeof (ncm) };
				SystemParametersInfo (SPI_GETNONCLIENTMETRICS, sizeof (ncm), &ncm, 0);
				LOGFONT lf = ncm.lfMessageFont;
				m_font.CreateFontIndirect (&lf);
				LOGFONT bold = lf;
				bold.lfWeight = FW_SEMIBOLD;
				m_bold.CreateFontIndirect (&bold);
				LOGFONT big = lf;
				big.lfHeight = lf.lfHeight * 3 / 2;
				big.lfWeight = FW_SEMIBOLD;
				m_big.CreateFontIndirect (&big);
				SetFont (&m_font);
				{
				CClientDC dc (this);
				TEXTMETRIC tm;
				CFont *old = dc.SelectObject (&m_font);
				dc.GetTextMetrics (&tm);
				m_u = (std::max) (12, static_cast<int> (tm.tmHeight));
				dc.SelectObject (&m_big);
				dc.GetTextMetrics (&tm);
				m_bigH = tm.tmHeight;
				dc.SelectObject (old);
				}

				auto label = [&] (CStatic &c, const std::wstring &text, int id, CFont &f)
					{
					// One line each; a long one ends in "..." rather than wrapping out of sight.
					c.Create (text.c_str (), WS_CHILD | WS_VISIBLE | SS_LEFTNOWORDWRAP | SS_NOPREFIX | SS_ENDELLIPSIS,
							  CRect (), this, id);
					c.SetFont (&f);
					};
				auto button = [&] (CButton &b, const wchar_t *text, int id, DWORD style)
					{
					b.Create (text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | style, CRect (), this, id);
					b.SetFont (&m_font);
					};
				label (m_title, L"Dump lathe parameters", IdTitle, m_big);
				label (m_summary, std::filesystem::path (m_part).filename ().wstring (), IdSummary, m_font);
				label (m_opsHead, L"Operations", IdOpsHead, m_bold);
				label (m_folderHead, L"Save to", IdFolderHead, m_bold);
				label (m_nameHead, L"File name", IdNameHead, m_bold);
				label (m_skippedCtl, m_skipped.empty () ? L"" : L"Not read yet (left out): " + m_skipped,
					   IdSkipped, m_font);
				label (m_preview, L"", IdPreview, m_bold);
				label (m_tokens, L"Tokens:  {part}  {date}  {time}  {scope}  {ops}    "
								 L"An existing file is never overwritten.", IdTokens, m_font);

				// ---- The operations: a tree of toolpath groups, a tick box each.
				int selectedInMgr = 0;
				for (const DumpDialog::Op &o : m_ops)
					selectedInMgr += o.selectedInMgr;
				button (m_allBtn, L"All", IdAll, BS_PUSHBUTTON);
				button (m_noneBtn, L"None", IdNone, BS_PUSHBUTTON);
				button (m_mgrBtn, (L"Selected in Operation Manager (" + std::to_wstring (selectedInMgr) + L")").c_str (),
						IdMgr, BS_PUSHBUTTON);
				m_mgrBtn.EnableWindow (selectedInMgr > 0);
				m_kind.Create (WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWNLIST, CRect (0, 0, 10, 300),
							   this, IdKind);
				m_kind.SetFont (&m_font);
				std::map<std::wstring, int> kinds;
				for (const DumpDialog::Op &o : m_ops)
					++kinds[o.kind];
				for (const auto &kv : kinds)
					{
					const int at = m_kind.AddString ((kv.first + L"  (" + std::to_wstring (kv.second) + L")").c_str ());
					m_kindNames.push_back (kv.first);
					m_kind.SetItemData (at, static_cast<DWORD_PTR> (m_kindNames.size () - 1));
					}
				m_kind.SetCurSel (0);
				button (m_kindOn, L"Tick kind", IdKindOn, BS_PUSHBUTTON);
				button (m_kindOff, L"Untick kind", IdKindOff, BS_PUSHBUTTON);

				// ---- Find: type part of a comment, tool, type or op number; what
				// matches is highlighted in the tree, the rest faded.
				label (m_findHead, L"Find", IdFindHead, m_font);
				m_findHead.ModifyStyle (0, SS_CENTERIMAGE);
				m_find.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, CRect (), this, IdFind);
				m_find.SetFont (&m_font);
				m_find.SetCueBanner (L"comment, tool, type or op number");
				button (m_findOn, L"Tick matches", IdFindOn, BS_PUSHBUTTON);
				button (m_findOff, L"Untick matches", IdFindOff, BS_PUSHBUTTON);
				label (m_findCount, L"", IdFindCount, m_font);
				m_findCount.ModifyStyle (0, SS_CENTERIMAGE);
				m_findOn.EnableWindow (FALSE);
				m_findOff.EnableWindow (FALSE);
				m_match.assign (m_ops.size (), 0);

				m_tree.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | TVS_HASBUTTONS | TVS_HASLINES
								   | TVS_LINESATROOT | TVS_SHOWSELALWAYS | TVS_FULLROWSELECT,
							   CRect (), this, IdTree);
				// Tick boxes have to be switched on after the tree exists.
				m_tree.ModifyStyle (0, TVS_CHECKBOXES);
				m_tree.SetFont (&m_font);
				m_filling = true;
				std::map<long, HTREEITEM> groupItem;
				for (size_t i = 0; i < m_ops.size (); ++i)
					{
					const DumpDialog::Op &o = m_ops[i];
					HTREEITEM g;
					const auto it = groupItem.find (o.group);
					if (it == groupItem.end ())
						{
						g = m_tree.InsertItem ((o.groupName.empty () ? L"(no group)" : o.groupName).c_str ());
						m_tree.SetItemData (g, static_cast<DWORD_PTR> (-1));
						groupItem[o.group] = g;
						m_groups.push_back (g);
						}
					else
						g = it->second;
					HTREEITEM h = m_tree.InsertItem (o.text.c_str (), g);
					m_tree.SetItemData (h, static_cast<DWORD_PTR> (i));
					m_items.push_back (h);
					}
				for (HTREEITEM g : m_groups)
					m_tree.Expand (g, TVE_EXPAND);
				for (size_t i = 0; i < m_ops.size (); ++i)
					m_tree.SetCheck (m_items[i], m_ops[i].on);
				SyncGroups ();
				m_filling = false;
				if (!m_items.empty ())
					m_tree.EnsureVisible (m_tree.GetRootItem ());

				// ---- Saved ticks: name the ticks, pick the name later to tick them
				// again. Kept per part file name.
				label (m_savedHead, L"Saved ticks", IdSavedHead, m_font);
				m_savedHead.ModifyStyle (0, SS_CENTERIMAGE);
				m_saved.Create (WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_DROPDOWN | CBS_AUTOHSCROLL,
								CRect (0, 0, 10, 300), this, IdSaved);
				m_saved.SetFont (&m_font);
				m_saved.SetCueBanner (L"name the ticks, or pick a name");
				button (m_save, L"Save ticks", IdSave, BS_PUSHBUTTON);
				button (m_delete, L"Delete", IdDelete, BS_PUSHBUTTON);
				label (m_savedNote, L"", IdSavedNote, m_font);
				m_savedNote.ModifyStyle (0, SS_CENTERIMAGE);
				FillSaved ();

				// ---- Where, what name, options.
				const std::wstring partFolder = std::filesystem::path (m_part).parent_path ().wstring ();
				m_folder.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, CRect (), this, IdFolder);
				m_folder.SetFont (&m_font);
				m_folder.SetWindowText ((m_settings.folder.empty () ? partFolder : m_settings.folder).c_str ());
				button (m_browse, L"Browse...", IdBrowse, BS_PUSHBUTTON);
				button (m_beside, L"Beside the part", IdBeside, BS_PUSHBUTTON);

				m_pattern.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, CRect (), this, IdPattern);
				m_pattern.SetFont (&m_font);
				m_pattern.SetWindowText ((m_settings.pattern.empty () ? FileRules::kDefaultPattern
																		: m_settings.pattern).c_str ());

				button (m_open, L"Open in Excel when done", IdOpen, BS_AUTOCHECKBOX);
				m_open.SetCheck (m_settings.openExcel ? BST_CHECKED : BST_UNCHECKED);
				button (m_pics, L"Tool pictures on the Tools sheet (lathe tools)", IdPictures, BS_AUTOCHECKBOX);
				m_pics.SetCheck (m_settings.pictures ? BST_CHECKED : BST_UNCHECKED);
				button (m_sim, L"Simulate the stock to measure what each op removes (lathe) - off: Mastercam's "
							   L"own stock boundaries", IdStockSim, BS_AUTOCHECKBOX);
				m_sim.SetCheck (m_settings.stockSim ? BST_CHECKED : BST_UNCHECKED);
				button (m_macros, L"Include macros (.xlsm): a Parameter Table ribbon tab - bulk edit, "
								  L"revert, change list, calculators", IdMacros, BS_AUTOCHECKBOX);
				m_macros.SetCheck (m_settings.macros ? BST_CHECKED : BST_UNCHECKED);

				button (m_ok, L"Dump", IDOK, BS_DEFPUSHBUTTON);
				m_ok.SetFont (&m_bold);
				button (m_cancel, L"Cancel", IDCANCEL, BS_PUSHBUTTON);

				// Size: wide enough for the tree and the path, tall enough for all of it.
				CRect work;
				SystemParametersInfo (SPI_GETWORKAREA, 0, &work, 0);
				const int rowH = m_tree.GetItemHeight ();
				const int rows = static_cast<int> ((std::min) (m_ops.size () + m_groups.size (), static_cast<size_t> (16)));
				m_treeH = (std::max) (6, rows) * rowH + rowH;
				// Wide enough for every row of buttons at this font - a large font needs more.
				const int w = (std::max) (m_u * 42, Needed ());
				int h = Layout (w, true);
				if (h > work.Height () * 9 / 10)
					{
					m_treeH -= h - work.Height () * 9 / 10;
					h = Layout (w, true);
					}
				CRect wr (0, 0, w, h);
				CalcWindowRect (&wr);
				SetWindowPos (nullptr, work.left + (work.Width () - wr.Width ()) / 2,
							  work.top + (work.Height () - wr.Height ()) / 2, wr.Width (), wr.Height (), SWP_NOZORDER);
				Layout (w, false);
				Update ();
				m_ok.SetFocus ();
				return FALSE;
				}

			/// Place everything for a client width; returns the height used.
			int Layout (int w, bool measureOnly)
				{
				const int u = m_u, pad = u, gap = u / 2, inner = w - 2 * pad;
				const int lineH = u * 3 / 2, editH = u * 3 / 2, btnH = u * 2;
				int y = pad;
				auto place = [&] (CWnd &c, int x, int yy, int cw, int ch)
					{
					if (!measureOnly)
						c.MoveWindow (x, yy, cw, ch);
					};
				place (m_title, pad, y, inner, m_bigH + 2);
				y += m_bigH + 2;
				place (m_summary, pad, y, inner, lineH);
				y += lineH + gap;

				place (m_opsHead, pad, y, inner, lineH);
				y += lineH;
				// Rows of buttons, each sized to its words:
				//   All | None | Selected in Operation Manager (n)
				//   Kind: [ROUGH (4)  v] Tick kind | Untick kind
				//   Find [..........] Tick matches | Untick matches   3 found
				//   (the tree)
				//   Saved ticks [name   v] Save ticks | Delete   Saved "finish"
				int x = pad + u;
				for (CButton *b : { &m_allBtn, &m_noneBtn, &m_mgrBtn })
					{
					const int bw = (std::max) (u * 4, ButtonW (*b));
					place (*b, x, y, bw, editH);
					x += bw + gap / 2;
					}
				y += editH + gap / 2;
				x = pad + u;
				const int kindW = KindW ();
				if (!measureOnly)
					m_kind.MoveWindow (x, y, kindW, editH + u * 12);
				x += kindW + gap / 2;
				for (CButton *b : { &m_kindOn, &m_kindOff })
					{
					const int bw = ButtonW (*b);
					place (*b, x, y, bw, editH);
					x += bw + gap / 2;
					}
				y += editH + gap / 2;

				// A row of: label, a box taking what is left, buttons, a note.
				auto row = [&] (CStatic &head, CWnd &box, std::initializer_list<CButton *> buttons, CStatic &note,
								int noteW, int boxH)
					{
					int bx = pad + u;
					const int hw = LabelW (head);
					place (head, bx, y, hw, editH);
					bx += hw + gap / 2;
					int right = w - pad - noteW;
					for (CButton *b : buttons)
						right -= ButtonW (*b) + gap / 2;
					const int boxW = (std::max) (u * 8, right - bx);
					if (!measureOnly)
						box.MoveWindow (bx, y, boxW, boxH);
					bx += boxW + gap / 2;
					for (CButton *b : buttons)
						{
						const int bw = ButtonW (*b);
						place (*b, bx, y, bw, editH);
						bx += bw + gap / 2;
						}
					place (note, bx, y, (std::max) (0, w - pad - bx), editH);
					y += editH + gap / 2;
					};
				row (m_findHead, m_find, { &m_findOn, &m_findOff }, m_findCount, FindNoteW (), editH);
				place (m_tree, pad + u, y, inner - u, m_treeH);
				y += m_treeH + gap / 2;
				row (m_savedHead, m_saved, { &m_save, &m_delete }, m_savedNote, u * 10, editH + u * 12);
				y -= gap / 2;
				if (!m_skipped.empty ())
					{
					place (m_skippedCtl, pad + u, y, inner - u, lineH);
					y += lineH;
					}
				y += gap;

				const int bw = u * 7, bw2 = u * 9;
				place (m_folderHead, pad, y, inner, lineH);
				y += lineH;
				place (m_folder, pad + u, y, inner - u - bw - bw2 - 2 * gap, editH);
				place (m_browse, w - pad - bw - bw2 - gap, y, bw, editH);
				place (m_beside, w - pad - bw2, y, bw2, editH);
				y += editH + gap;

				place (m_nameHead, pad, y, inner, lineH);
				y += lineH;
				place (m_pattern, pad + u, y, inner - u, editH);
				y += editH + gap / 2;
				place (m_preview, pad + u, y, inner - u, lineH);
				y += lineH;
				place (m_tokens, pad + u, y, inner - u, lineH);
				y += lineH + gap;

				for (CButton *b : { &m_open, &m_pics, &m_sim, &m_macros })
					{
					place (*b, pad, y, inner, lineH);
					y += lineH;
					}
				y += gap;

				const int okW = u * 13, cW = u * 7;
				place (m_ok, w - pad - okW - gap - cW, y, okW, btnH);
				place (m_cancel, w - pad - cW, y, cW, btnH);
				y += btnH + pad;
				return y;
				}

			int TextW (const CString &text)
				{
				CClientDC dc (this);
				CFont *old = dc.SelectObject (&m_font);
				const int cx = dc.GetTextExtent (text).cx;
				dc.SelectObject (old);
				return cx;
				}

			/// A button sized to its words.
			int ButtonW (CWnd &c)
				{
				CString text;
				c.GetWindowText (text);
				return TextW (text) + m_u * 2;
				}

			/// A label sized to its words.
			int LabelW (CWnd &c)
				{
				CString text;
				c.GetWindowText (text);
				return TextW (text) + m_u / 2;
				}

			/// The kind list: as wide as its longest entry.
			int KindW ()
				{
				int kindW = m_u * 8;
				for (int i = 0; i < m_kind.GetCount (); ++i)
					{
					CString item;
					m_kind.GetLBText (i, item);
					kindW = (std::max) (kindW, TextW (item) + m_u * 3);
					}
				return kindW;
				}

			/// Room for the search's count, at its longest.
			int FindNoteW ()
				{
				return TextW (L"none found") + m_u;
				}

			/// The client width every row of controls needs, with a box at its
			/// smallest - so a large font widens the window instead of overlapping.
			int Needed ()
				{
				const int u = m_u, gap = u / 2, edge = 2 * u + u;		// padding both sides + the indent
				int row1 = 0;
				for (CButton *b : { &m_allBtn, &m_noneBtn, &m_mgrBtn })
					row1 += (std::max) (u * 4, ButtonW (*b)) + gap / 2;
				const int row2 = KindW () + ButtonW (m_kindOn) + ButtonW (m_kindOff) + gap * 3 / 2;
				const int row3 = LabelW (m_findHead) + u * 12 + ButtonW (m_findOn) + ButtonW (m_findOff) + FindNoteW () + gap * 2;
				const int row4 = LabelW (m_savedHead) + u * 12 + ButtonW (m_save) + ButtonW (m_delete) + u * 10 + gap * 2;
				return edge + (std::max) ({ row1, row2, row3, row4 });
				}

			int Count () const
				{
				int n = 0;
				for (HTREEITEM h : m_items)
					n += m_tree.GetCheck (h) ? 1 : 0;
				return n;
				}

			bool AllSelectedInMgr () const
				{
				int sel = 0, ticked = 0, match = 0;
				for (size_t i = 0; i < m_items.size (); ++i)
					{
					const bool t = m_tree.GetCheck (m_items[i]) != FALSE;
					sel += m_ops[i].selectedInMgr;
					ticked += t;
					match += t && m_ops[i].selectedInMgr;
					}
				return sel > 0 && ticked == sel && match == sel;
				}

			/// A group is ticked when all of its operations are.
			void SyncGroups ()
				{
				for (HTREEITEM g : m_groups)
					{
					bool all = true;
					for (HTREEITEM c = m_tree.GetChildItem (g); c != nullptr; c = m_tree.GetNextSiblingItem (c))
						all = all && m_tree.GetCheck (c);
					m_tree.SetCheck (g, all);
					}
				}

			void SetWhere (bool (*want) (const DumpDialog::Op &, const std::wstring &), const std::wstring &arg, bool on)
				{
				m_filling = true;
				for (size_t i = 0; i < m_items.size (); ++i)
					if (want (m_ops[i], arg))
						m_tree.SetCheck (m_items[i], on);
				SyncGroups ();
				m_filling = false;
				Update ();
				}

			/// The button and the file name preview follow every change.
			void Update ()
				{
				if (m_ok.GetSafeHwnd () == nullptr || m_filling)
					return;
				const int n = Count ();
				m_ok.SetWindowText (n == 1 ? L"Dump 1 operation"
										   : (L"Dump " + std::to_wstring (n) + L" operations").c_str ());
				m_ok.EnableWindow (n > 0);

				CString pat, folder;
				m_pattern.GetWindowText (pat);
				m_folder.GetWindowText (folder);
				const std::wstring name = FileRules::Name (pat.GetString (),
														   std::filesystem::path (m_part).stem ().wstring (),
														   std::time (nullptr), n < static_cast<int> (m_items.size ()),
														   static_cast<size_t> (n),
														   m_macros.GetCheck () == BST_CHECKED ? L".xlsm" : L".xlsx");
				std::error_code ec;
				const bool folderOk = std::filesystem::is_directory (folder.GetString (), ec);
				m_previewOk = folderOk;
				m_preview.SetWindowText (folderOk ? (L"→  " + name).c_str ()
												  : L"→  that folder does not exist");
				m_preview.Invalidate ();
				}

			void OnOK () override
				{
				CString folder, pat;
				m_folder.GetWindowText (folder);
				m_pattern.GetWindowText (pat);
				std::error_code ec;
				if (!std::filesystem::is_directory (folder.GetString (), ec))
					{
					MessageBoxW (L"That folder does not exist.", L"Parameter Table Tool", MB_ICONWARNING);
					return;
					}
				const std::wstring partFolder = std::filesystem::path (m_part).parent_path ().wstring ();
				m_settings.folder = _wcsicmp (folder.GetString (), partFolder.c_str ()) == 0 ? L"" : folder.GetString ();
				m_settings.pattern = pat.GetString ();
				m_settings.openExcel = m_open.GetCheck () == BST_CHECKED;
				m_settings.pictures = m_pics.GetCheck () == BST_CHECKED;
				m_settings.macros = m_macros.GetCheck () == BST_CHECKED;
				m_settings.stockSim = m_sim.GetCheck () == BST_CHECKED;
				for (size_t i = 0; i < m_items.size (); ++i)
					m_ops[i].on = m_tree.GetCheck (m_items[i]) != FALSE;
				CDialog::OnOK ();
				}

			afx_msg void OnAll ()
				{
				SetWhere ([] (const DumpDialog::Op &, const std::wstring &) { return true; }, L"", true);
				}

			afx_msg void OnNone ()
				{
				SetWhere ([] (const DumpDialog::Op &, const std::wstring &) { return true; }, L"", false);
				}

			afx_msg void OnMgr ()
				{
				m_filling = true;
				for (size_t i = 0; i < m_items.size (); ++i)
					m_tree.SetCheck (m_items[i], m_ops[i].selectedInMgr);
				SyncGroups ();
				m_filling = false;
				Update ();
				}

			afx_msg void OnKind (bool on)
				{
				const int sel = m_kind.GetCurSel ();
				if (sel < 0)
					return;
				const std::wstring kind = m_kindNames[static_cast<size_t> (m_kind.GetItemData (sel))];
				SetWhere ([] (const DumpDialog::Op &o, const std::wstring &k) { return o.kind == k; }, kind, on);
				}
			afx_msg void OnKindOn () { OnKind (true); }
			afx_msg void OnKindOff () { OnKind (false); }

			/// A tick changed: a group passes it to its operations; an operation's
			/// group follows its operations.
			afx_msg void OnTreeChanged (NMHDR *nm, LRESULT *result)
				{
				*result = 0;
				if (m_filling)
					return;
				const NMTVITEMCHANGE *c = reinterpret_cast<const NMTVITEMCHANGE *> (nm);
				if (((c->uStateNew ^ c->uStateOld) & TVIS_STATEIMAGEMASK) == 0)
					return;
				m_filling = true;
				const bool on = m_tree.GetCheck (c->hItem) != FALSE;
				if (m_tree.GetItemData (c->hItem) == static_cast<DWORD_PTR> (-1))
					for (HTREEITEM k = m_tree.GetChildItem (c->hItem); k != nullptr; k = m_tree.GetNextSiblingItem (k))
						m_tree.SetCheck (k, on);
				SyncGroups ();
				m_filling = false;
				Update ();
				}

			// ---- Find.

			/// What is typed: every operation it matches is highlighted (the tree
			/// draws it), the first one scrolled into view, and counted.
			afx_msg void OnFind ()
				{
				CString q;
				m_find.GetWindowText (q);
				m_query = q.GetString ();
				m_searching = m_query.find_first_not_of (L" \t") != std::wstring::npos;
				int n = 0;
				HTREEITEM first = nullptr;
				for (size_t i = 0; i < m_ops.size () && i < m_items.size (); ++i)
					{
					m_match[i] = Pick::Matches (m_ops[i].text, m_query);
					n += m_match[i];
					if (m_match[i] && first == nullptr)
						first = m_items[i];
					}
				m_findCount.SetWindowText (!m_searching ? L"" : n == 0 ? L"none found"
																	   : (std::to_wstring (n) + L" found").c_str ());
				m_findOn.EnableWindow (n > 0);
				m_findOff.EnableWindow (n > 0);
				if (first != nullptr)
					m_tree.EnsureVisible (first);
				m_tree.Invalidate ();
				}

			void TickMatches (bool on)
				{
				SetWhere ([] (const DumpDialog::Op &o, const std::wstring &q) { return Pick::Matches (o.text, q); },
						  m_query, on);
				}
			afx_msg void OnFindOn () { TickMatches (true); }
			afx_msg void OnFindOff () { TickMatches (false); }

			/// While searching: a match drawn on a highlight, the rest faded. The
			/// selected row and the group rows keep their own look.
			afx_msg void OnTreeDraw (NMHDR *nm, LRESULT *result)
				{
				NMTVCUSTOMDRAW *cd = reinterpret_cast<NMTVCUSTOMDRAW *> (nm);
				*result = CDRF_DODEFAULT;
				if (cd->nmcd.dwDrawStage == CDDS_PREPAINT)
					{
					*result = CDRF_NOTIFYITEMDRAW;
					return;
					}
				if (cd->nmcd.dwDrawStage != CDDS_ITEMPREPAINT || !m_searching || (cd->nmcd.uItemState & CDIS_SELECTED))
					return;
				const DWORD_PTR i = cd->nmcd.lItemlParam;
				if (i == static_cast<DWORD_PTR> (-1) || i >= m_match.size ())
					return;
				if (m_match[i])
					{
					cd->clrText = kInk;
					cd->clrTextBk = kFound;
					}
				else
					cd->clrText = kFaded;
				*result = CDRF_NEWFONT;
				}

			// ---- Saved ticks.

			void FillSaved ()
				{
				m_selections = Settings::Selections (m_part);
				m_saved.ResetContent ();
				for (const auto &sel : m_selections)
					m_saved.AddString (sel.first.c_str ());
				}

			static std::wstring Ops (size_t n)
				{
				return std::to_wstring (n) + (n == 1 ? L" op" : L" ops");
				}

			void Note (const std::wstring &text)
				{
				m_savedNote.SetWindowText (text.c_str ());
				}

			/// A name picked from the list: its operations ticked, the rest not.
			afx_msg void OnSavedPick ()
				{
				const int sel = m_saved.GetCurSel ();
				if (sel < 0 || static_cast<size_t> (sel) >= m_selections.size ())
					return;
				const auto &pick = m_selections[static_cast<size_t> (sel)];
				const std::set<long> ids (pick.second.begin (), pick.second.end ());
				std::set<long> found;
				m_filling = true;
				for (size_t i = 0; i < m_items.size (); ++i)
					{
					const bool on = ids.count (m_ops[i].idn) > 0;
					m_tree.SetCheck (m_items[i], on);
					if (on)
						found.insert (m_ops[i].idn);
					}
				SyncGroups ();
				m_filling = false;
				Update ();
				const size_t gone = ids.size () - found.size ();
				Note (L"Ticked " + Ops (found.size ()) + (gone > 0 ? L", " + std::to_wstring (gone) + L" gone from the part"
																: std::wstring ()));
				}

			/// The ticks saved under the name typed (the same name is replaced).
			afx_msg void OnSave ()
				{
				CString t;
				m_saved.GetWindowText (t);
				const std::wstring name = Pick::CleanName (t.GetString ());
				if (name.empty ())
					{
					Note (L"Type a name first");
					m_saved.SetFocus ();
					return;
					}
				std::vector<long> ids;
				for (size_t i = 0; i < m_items.size (); ++i)
					if (m_tree.GetCheck (m_items[i]))
						ids.push_back (m_ops[i].idn);
				if (ids.empty ())
					{
					Note (L"Nothing ticked");
					return;
					}
				Settings::SaveSelection (m_part, name, ids);
				FillSaved ();
				m_saved.SetWindowText (name.c_str ());
				Note (L"Saved " + Ops (ids.size ()));
				}

			afx_msg void OnDelete ()
				{
				CString t;
				m_saved.GetWindowText (t);
				const std::wstring name = Pick::CleanName (t.GetString ());
				for (const auto &sel : m_selections)
					if (_wcsicmp (sel.first.c_str (), name.c_str ()) == 0)
						{
						Settings::DeleteSelection (m_part, sel.first);
						FillSaved ();
						m_saved.SetWindowText (L"");
						Note (L"Deleted");
						return;
						}
				Note (name.empty () ? L"Pick a name first" : L"No saved ticks by that name");
				}

			/// Enter in the find box ticks what it found; in the name box it saves -
			/// never the window's Dump. Escape in the find box clears the search first.
			BOOL PreTranslateMessage (MSG *msg) override
				{
				if (msg->message == WM_KEYDOWN && (msg->wParam == VK_RETURN || msg->wParam == VK_ESCAPE))
					{
					const HWND focus = ::GetFocus ();
					if (focus != nullptr && focus == m_find.GetSafeHwnd ())
						{
						if (msg->wParam == VK_ESCAPE && m_find.GetWindowTextLength () > 0)
							{
							m_find.SetWindowText (L"");
							return TRUE;
							}
						if (msg->wParam == VK_RETURN)
							{
							if (m_findOn.IsWindowEnabled ())
								OnFindOn ();
							return TRUE;
							}
						}
					else if (focus != nullptr && msg->wParam == VK_RETURN && ::GetParent (focus) == m_saved.GetSafeHwnd ()
							 && !m_saved.GetDroppedState ())
						{
						OnSave ();
						return TRUE;
						}
					}
				return CDialog::PreTranslateMessage (msg);
				}

			afx_msg void OnBrowse ()
				{
				CString folder;
				m_folder.GetWindowText (folder);
				CFolderPickerDialog dlg (folder, 0, this);
				if (dlg.DoModal () == IDOK)
					m_folder.SetWindowText (dlg.GetPathName ());
				Update ();
				}

			afx_msg void OnBeside ()
				{
				m_folder.SetWindowText (std::filesystem::path (m_part).parent_path ().wstring ().c_str ());
				Update ();
				}

			afx_msg void OnChange () { Update (); }

			afx_msg HBRUSH OnCtlColor (CDC *dc, CWnd *wnd, UINT ctl)
				{
				HBRUSH br = CDialog::OnCtlColor (dc, wnd, ctl);
				const int id = wnd != nullptr ? wnd->GetDlgCtrlID () : 0;
				if (id == IdSummary || id == IdTokens || id == IdSkipped || id == IdFindCount || id == IdSavedNote)
					dc->SetTextColor (kMuted);
				else if (id == IdPreview)
					dc->SetTextColor (m_previewOk ? kGood : RGB (0xB4, 0x23, 0x18));
				else if (id == IdTitle || id == IdOpsHead || id == IdFolderHead || id == IdNameHead)
					dc->SetTextColor (kInk);
				return br;
				}

			DECLARE_MESSAGE_MAP ()

		private:
			std::wstring m_part;
			std::vector<DumpDialog::Op> &m_ops;
			std::wstring m_skipped;
			Settings::Dump &m_settings;
			CWnd *m_parent = nullptr;
			std::vector<WORD> m_tpl;
			int m_u = 16, m_bigH = 24, m_treeH = 200;
				bool m_filling = false, m_previewOk = true;
			std::wstring m_query;					// what the find box holds
			bool m_searching = false;				// it holds more than spaces
			std::vector<char> m_match;				// per op: the search found it
			std::vector<std::pair<std::wstring, std::vector<long>>> m_selections;	// the part's saved ticks, as listed

			std::vector<HTREEITEM> m_items;		// one per op, as m_ops
			std::vector<HTREEITEM> m_groups;
			std::vector<std::wstring> m_kindNames;

			CFont m_font, m_bold, m_big;
			CStatic m_title, m_summary, m_opsHead, m_folderHead, m_nameHead, m_skippedCtl, m_preview, m_tokens;
			CButton m_allBtn, m_noneBtn, m_mgrBtn, m_kindOn, m_kindOff;
				CButton m_browse, m_beside, m_open, m_pics, m_sim, m_macros, m_ok, m_cancel;
			CStatic m_findHead, m_findCount, m_savedHead, m_savedNote;
			CButton m_findOn, m_findOff, m_save, m_delete;
			CComboBox m_kind, m_saved;
			CTreeCtrl m_tree;
			CEdit m_folder, m_pattern, m_find;
		};

	BEGIN_MESSAGE_MAP (Dlg, CDialog)
		ON_BN_CLICKED (IdBrowse, &Dlg::OnBrowse)
		ON_BN_CLICKED (IdBeside, &Dlg::OnBeside)
		ON_BN_CLICKED (IdAll, &Dlg::OnAll)
		ON_BN_CLICKED (IdNone, &Dlg::OnNone)
		ON_BN_CLICKED (IdMgr, &Dlg::OnMgr)
		ON_BN_CLICKED (IdKindOn, &Dlg::OnKindOn)
		ON_BN_CLICKED (IdKindOff, &Dlg::OnKindOff)
		ON_EN_CHANGE (IdPattern, &Dlg::OnChange)
		ON_EN_CHANGE (IdFolder, &Dlg::OnChange)
		ON_BN_CLICKED (IdMacros, &Dlg::OnChange)
		ON_NOTIFY (TVN_ITEMCHANGED, IdTree, &Dlg::OnTreeChanged)
		ON_NOTIFY (NM_CUSTOMDRAW, IdTree, &Dlg::OnTreeDraw)
		ON_EN_CHANGE (IdFind, &Dlg::OnFind)
		ON_BN_CLICKED (IdFindOn, &Dlg::OnFindOn)
		ON_BN_CLICKED (IdFindOff, &Dlg::OnFindOff)
		ON_CBN_SELCHANGE (IdSaved, &Dlg::OnSavedPick)
		ON_BN_CLICKED (IdSave, &Dlg::OnSave)
		ON_BN_CLICKED (IdDelete, &Dlg::OnDelete)
		ON_WM_CTLCOLOR ()
	END_MESSAGE_MAP ()
	}

namespace DumpDialog
	{
	bool Show (const std::wstring &partFile, std::vector<Op> &ops, const std::wstring &skipped,
			   Settings::Dump &settings)
		{
		Dlg dlg (partFile, ops, skipped, settings, CWnd::FromHandle (get_MainFrame ()->GetSafeHwnd ()));
		return dlg.Run () == IDOK;
		}
	}
