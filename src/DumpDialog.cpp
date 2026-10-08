#include "stdafx.h"
#include "MastercamSdk.h"
#include "DumpDialog.h"
#include "FileRules.h"

#include <algorithm>
#include <ctime>
#include <filesystem>

namespace
	{
	const COLORREF kInk   = RGB (0x1F, 0x29, 0x37);
	const COLORREF kMuted = RGB (0x5B, 0x65, 0x73);
	const COLORREF kGood  = RGB (0x06, 0x76, 0x47);

	enum
		{
		IdTitle = 2001, IdSummary, IdOpsHead, IdAll, IdSelected, IdKinds, IdSkipped,
		IdFolderHead, IdFolder, IdBrowse, IdBeside, IdNameHead, IdPattern, IdPreview,
		IdTokens, IdOpen, IdPictures
		};

	class Dlg : public CDialog
		{
		public:
			Dlg (const std::wstring &part, std::vector<DumpDialog::Kind> &kinds, const std::wstring &skipped,
				 bool &selectedOnly, Settings::Dump &settings, CWnd *parent)
				: m_part (part), m_kinds (kinds), m_skipped (skipped), m_selectedOnly (selectedOnly),
				  m_settings (settings), m_parent (parent)
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

				int all = 0, sel = 0;
				for (const DumpDialog::Kind &k : m_kinds)
					{
					all += k.all;
					sel += k.selected;
					}
				m_all.Create ((L"All operations in the part  (" + std::to_wstring (all) + L")").c_str (),
							  WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_GROUP | BS_AUTORADIOBUTTON, CRect (), this, IdAll);
				m_sel.Create ((L"Only the ones selected in the Operation Manager  (" + std::to_wstring (sel) + L")").c_str (),
							  WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTORADIOBUTTON, CRect (), this, IdSelected);
				m_all.SetFont (&m_font);
				m_sel.SetFont (&m_font);
				if (sel == 0)
					{
					m_sel.EnableWindow (FALSE);
					m_selectedOnly = false;
					}
				(m_selectedOnly ? m_sel : m_all).SetCheck (BST_CHECKED);

				m_list.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL
								   | LVS_NOSORTHEADER, CRect (), this, IdKinds);
				m_list.SetExtendedStyle (LVS_EX_CHECKBOXES | LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);
				m_list.SetFont (&m_font);
				m_rowImages.Create (1, m_u * 3 / 2, ILC_COLOR, 1, 0);
				m_list.SetImageList (&m_rowImages, LVSIL_SMALL);
				m_list.InsertColumn (0, L"Kind", LVCFMT_LEFT, m_u * 12);
				m_list.InsertColumn (1, L"In the part", LVCFMT_RIGHT, m_u * 6);
				m_list.InsertColumn (2, L"Selected", LVCFMT_RIGHT, m_u * 6);
				m_filling = true;
				for (size_t i = 0; i < m_kinds.size (); ++i)
					{
					const int at = m_list.InsertItem (static_cast<int> (i), m_kinds[i].name.c_str ());
					m_list.SetItemText (at, 1, std::to_wstring (m_kinds[i].all).c_str ());
					m_list.SetItemText (at, 2, std::to_wstring (m_kinds[i].selected).c_str ());
					m_list.SetCheck (at, m_kinds[i].on);
					}
				m_filling = false;

				const std::wstring partFolder = std::filesystem::path (m_part).parent_path ().wstring ();
				m_folder.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, CRect (), this, IdFolder);
				m_folder.SetFont (&m_font);
				m_folder.SetWindowText ((m_settings.folder.empty () ? partFolder : m_settings.folder).c_str ());
				m_browse.Create (L"Browse...", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, CRect (), this, IdBrowse);
				m_browse.SetFont (&m_font);
				m_beside.Create (L"Beside the part", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, CRect (), this, IdBeside);
				m_beside.SetFont (&m_font);

				m_pattern.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, CRect (), this, IdPattern);
				m_pattern.SetFont (&m_font);
				m_pattern.SetWindowText ((m_settings.pattern.empty () ? FileRules::kDefaultPattern
																		: m_settings.pattern).c_str ());

				m_open.Create (L"Open in Excel when done", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
							   CRect (), this, IdOpen);
				m_open.SetFont (&m_font);
				m_open.SetCheck (m_settings.openExcel ? BST_CHECKED : BST_UNCHECKED);
				m_pics.Create (L"Add a Tools sheet with tool pictures (lathe tools)",
							   WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, CRect (), this, IdPictures);
				m_pics.SetFont (&m_font);
				m_pics.SetCheck (m_settings.pictures ? BST_CHECKED : BST_UNCHECKED);

				m_ok.Create (L"Dump", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON, CRect (), this, IDOK);
				m_ok.SetFont (&m_bold);
				m_cancel.Create (L"Cancel", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, CRect (), this, IDCANCEL);
				m_cancel.SetFont (&m_font);

				// Size: wide enough for the list and the path, tall enough for all of it.
				const int rows = static_cast<int> ((std::min) (m_kinds.size (), static_cast<size_t> (9)));
				m_listH = (rows + 1) * (m_u * 3 / 2) + m_u;
				const int w = m_u * 38;
				const int h = Layout (w, true);
				CRect wr (0, 0, w, h);
				CalcWindowRect (&wr);
				CRect work;
				SystemParametersInfo (SPI_GETWORKAREA, 0, &work, 0);
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
				place (m_all, pad + u, y, inner - u, lineH);
				y += lineH;
				place (m_sel, pad + u, y, inner - u, lineH);
				y += lineH + gap / 2;
				place (m_list, pad + u, y, inner - u, m_listH);
				y += m_listH + gap / 2;
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

				place (m_open, pad, y, inner, lineH);
				y += lineH;
				place (m_pics, pad, y, inner, lineH);
				y += lineH + gap;

				const int okW = u * 13, cW = u * 7;
				place (m_ok, w - pad - okW - gap - cW, y, okW, btnH);
				place (m_cancel, w - pad - cW, y, cW, btnH);
				y += btnH + pad;
				return y;
				}

			int Count () const
				{
				int n = 0;
				for (int i = 0; i < m_list.GetItemCount (); ++i)
					if (m_list.GetCheck (i))
						n += m_sel.GetCheck () == BST_CHECKED ? m_kinds[static_cast<size_t> (i)].selected
															  : m_kinds[static_cast<size_t> (i)].all;
				return n;
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
														   std::time (nullptr), m_sel.GetCheck () == BST_CHECKED,
														   static_cast<size_t> (n));
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
				m_settings.skipKinds.clear ();
				for (int i = 0; i < m_list.GetItemCount (); ++i)
					{
					m_kinds[static_cast<size_t> (i)].on = m_list.GetCheck (i) != FALSE;
					if (!m_kinds[static_cast<size_t> (i)].on)
						m_settings.skipKinds += L"|" + m_kinds[static_cast<size_t> (i)].name;
					}
				if (!m_settings.skipKinds.empty ())
					m_settings.skipKinds += L"|";
				m_selectedOnly = m_sel.GetCheck () == BST_CHECKED;
				CDialog::OnOK ();
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

			afx_msg void OnItemChanged (NMHDR *, LRESULT *result)
				{
				*result = 0;
				Update ();
				}

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
			std::vector<DumpDialog::Kind> &m_kinds;
			std::wstring m_skipped;
			bool &m_selectedOnly;
			Settings::Dump &m_settings;
			CWnd *m_parent = nullptr;
			std::vector<WORD> m_tpl;
			int m_u = 16, m_bigH = 24, m_listH = 100;
			bool m_filling = false, m_previewOk = true;

			CFont m_font, m_bold, m_big;
			CStatic m_title, m_summary, m_opsHead, m_folderHead, m_nameHead, m_skippedCtl, m_preview, m_tokens;
			CButton m_all, m_sel, m_browse, m_beside, m_open, m_pics, m_ok, m_cancel;
			CListCtrl m_list;
			CImageList m_rowImages;
			CEdit m_folder, m_pattern;
		};

	BEGIN_MESSAGE_MAP (Dlg, CDialog)
		ON_BN_CLICKED (IdBrowse, &Dlg::OnBrowse)
		ON_BN_CLICKED (IdBeside, &Dlg::OnBeside)
		ON_BN_CLICKED (IdAll, &Dlg::OnChange)
		ON_BN_CLICKED (IdSelected, &Dlg::OnChange)
		ON_EN_CHANGE (IdPattern, &Dlg::OnChange)
		ON_EN_CHANGE (IdFolder, &Dlg::OnChange)
		ON_NOTIFY (LVN_ITEMCHANGED, IdKinds, &Dlg::OnItemChanged)
		ON_WM_CTLCOLOR ()
	END_MESSAGE_MAP ()
	}

namespace DumpDialog
	{
	bool Show (const std::wstring &partFile, std::vector<Kind> &kinds, const std::wstring &skipped,
			   bool &selectedOnly, Settings::Dump &settings)
		{
		Dlg dlg (partFile, kinds, skipped, selectedOnly, settings,
				 CWnd::FromHandle (get_MainFrame ()->GetSafeHwnd ()));
		return dlg.Run () == IDOK;
		}
	}
