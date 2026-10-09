#include "stdafx.h"
#include "MastercamSdk.h"
#include "DumpDialog.h"
#include "FileRules.h"

#include <algorithm>
#include <ctime>
#include <filesystem>
#include <map>

namespace
	{
	const COLORREF kInk   = RGB (0x1F, 0x29, 0x37);
	const COLORREF kMuted = RGB (0x5B, 0x65, 0x73);
	const COLORREF kGood  = RGB (0x06, 0x76, 0x47);

	enum
		{
		IdTitle = 2001, IdSummary, IdOpsHead, IdTree, IdAll, IdNone, IdMgr, IdKind, IdKindOn, IdKindOff,
		IdSkipped, IdFolderHead, IdFolder, IdBrowse, IdBeside, IdNameHead, IdPattern, IdPreview,
		IdTokens, IdOpen, IdPictures, IdMacros, IdStockSim
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
				const int w = m_u * 42;
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
				// Two rows of buttons, each sized to its words:
				//   All | None | Selected in Operation Manager (n)
				//   Kind: [ROUGH (4)  v] Tick kind | Untick kind
				auto widthOf = [&] (CWnd &c)
					{
					CString text;
					c.GetWindowText (text);
					CClientDC dc (this);
					CFont *old = dc.SelectObject (&m_font);
					const int cx = dc.GetTextExtent (text).cx;
					dc.SelectObject (old);
					return cx + u * 2;
					};
				int x = pad + u;
				for (CButton *b : { &m_allBtn, &m_noneBtn, &m_mgrBtn })
					{
					const int bw = (std::max) (u * 4, widthOf (*b));
					place (*b, x, y, bw, editH);
					x += bw + gap / 2;
					}
				y += editH + gap / 2;
				x = pad + u;
				int kindW = u * 8;
				for (int i = 0; i < m_kind.GetCount (); ++i)
					{
					CString item;
					m_kind.GetLBText (i, item);
					CClientDC dc (this);
					CFont *old = dc.SelectObject (&m_font);
					kindW = (std::max) (kindW, static_cast<int> (dc.GetTextExtent (item).cx) + u * 3);
					dc.SelectObject (old);
					}
				if (!measureOnly)
					m_kind.MoveWindow (x, y, kindW, editH + u * 12);
				x += kindW + gap / 2;
				for (CButton *b : { &m_kindOn, &m_kindOff })
					{
					const int bw = widthOf (*b);
					place (*b, x, y, bw, editH);
					x += bw + gap / 2;
					}
				y += editH + gap / 2;
				place (m_tree, pad + u, y, inner - u, m_treeH);
				y += m_treeH + gap / 2;
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
				if (id == IdSummary || id == IdTokens || id == IdSkipped)
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

			std::vector<HTREEITEM> m_items;		// one per op, as m_ops
			std::vector<HTREEITEM> m_groups;
			std::vector<std::wstring> m_kindNames;

			CFont m_font, m_bold, m_big;
			CStatic m_title, m_summary, m_opsHead, m_folderHead, m_nameHead, m_skippedCtl, m_preview, m_tokens;
			CButton m_allBtn, m_noneBtn, m_mgrBtn, m_kindOn, m_kindOff;
			CButton m_browse, m_beside, m_open, m_pics, m_sim, m_macros, m_ok, m_cancel;
			CComboBox m_kind;
			CTreeCtrl m_tree;
			CEdit m_folder, m_pattern;
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
