#include "stdafx.h"
#include "MastercamSdk.h"
#include "Preview.h"

#include <algorithm>

namespace
	{
	// ---- Colours: light, quiet, one meaning each.
	const COLORREF kInk        = RGB (0x1F, 0x29, 0x37);
	const COLORREF kMuted      = RGB (0x5B, 0x65, 0x73);
	const COLORREF kSectionBg  = RGB (0xE6, 0xEB, 0xF2);
	const COLORREF kOpBg       = RGB (0xF4, 0xF6, 0xFA);
	const COLORREF kOldInk     = RGB (0xB4, 0x23, 0x18);
	const COLORREF kOldBg      = RGB (0xFD, 0xEE, 0xEC);
	const COLORREF kNewInk     = RGB (0x06, 0x76, 0x47);
	const COLORREF kNewBg      = RGB (0xE8, 0xF6, 0xEE);
	const COLORREF kRefusedInk = RGB (0x8A, 0x4B, 0x00);
	const COLORREF kRefusedBg  = RGB (0xFF, 0xF5, 0xE1);
	const COLORREF kWhite      = RGB (0xFF, 0xFF, 0xFF);

	enum { IdList = 1001, IdTitle = 1002, IdSummary = 1003, IdFoot = 1004 };

	/// A value as one readable line: line breaks shown, nothing shown as such.
	std::wstring Show1 (const std::wstring &v)
		{
		if (v.empty ())
			return L"(empty)";
		std::wstring o;
		for (size_t i = 0; i < v.size (); ++i)
			{
			if (v[i] == L'\r')
				continue;
			if (v[i] == L'\n')
				o += L" ↵ ";
			else
				o += v[i];
			}
		return o;
		}

	class PreviewDlg : public CDialog
		{
		public:
			PreviewDlg (const std::wstring &title, const std::wstring &summary,
						const std::vector<Preview::Line> &lines, int changes, CWnd *parent)
				: m_title (title), m_summary (summary), m_lines (lines), m_changes (changes),
				  m_parent (parent)
				{
				}

			INT_PTR Run ()
				{
				// An empty template: every control is made in OnInitDialog, so the
				// window needs no resource and lays itself out when resized.
				m_tpl.assign (64, 0);
				DLGTEMPLATE *t = reinterpret_cast<DLGTEMPLATE *> (m_tpl.data ());
				t->style = WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MAXIMIZEBOX
						   | DS_MODALFRAME | DS_CENTER;
				t->cdit = 0;
				t->cx = 520;
				t->cy = 330;
				// menu, class and title are three empty WORDs after the header
				InitModalIndirect (t, m_parent);
				return DoModal ();
				}

		protected:
			BOOL OnInitDialog () override
				{
				CDialog::OnInitDialog ();
				SetWindowText (L"Parameter Table Tool - load preview");

				NONCLIENTMETRICS ncm = { sizeof (ncm) };
				SystemParametersInfo (SPI_GETNONCLIENTMETRICS, sizeof (ncm), &ncm, 0);
				LOGFONT lf = ncm.lfMessageFont;
				m_font.CreateFontIndirect (&lf);
				LOGFONT bold = lf;
				bold.lfWeight = FW_SEMIBOLD;
				m_bold.CreateFontIndirect (&bold);
				LOGFONT strike = lf;
				strike.lfStrikeOut = TRUE;
				m_strike.CreateFontIndirect (&strike);
				LOGFONT big = lf;
				big.lfHeight = lf.lfHeight * 3 / 2;
				big.lfWeight = FW_SEMIBOLD;
				m_big.CreateFontIndirect (&big);
				SetFont (&m_font);

				// EVERY SIZE FROM THE FONT. Mastercam runs scaled (150%, 175% ...),
				// and its fonts scale with it; a layout in fixed pixels then clips.
				// One unit = the height of a line of the dialog's text.
				{
				CClientDC dc (this);
				TEXTMETRIC tm;
				CFont *old = dc.SelectObject (&m_font);
				dc.GetTextMetrics (&tm);
				m_u = (std::max) (12, static_cast<int> (tm.tmHeight));
				dc.SelectObject (&m_big);
				dc.GetTextMetrics (&tm);
				m_bigH = tm.tmHeight;
				dc.SelectObject (&m_bold);
				const std::wstring longest = L"Apply 9999 changes";
				m_btnW = dc.GetTextExtent (longest.c_str (), static_cast<int> (longest.size ())).cx
						 + 2 * m_u;
				dc.SelectObject (old);
				}

				m_titleCtl.Create (m_title.c_str (), WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,
								   CRect (), this, IdTitle);
				m_titleCtl.SetFont (&m_big);
				m_summaryCtl.Create (m_summary.c_str (), WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,
									 CRect (), this, IdSummary);
				m_summaryCtl.SetFont (&m_font);
				m_footCtl.Create (m_changes > 0
									  ? L"There is no undo: every old value is written to ParamTable.log "
										L"first. Changed operations are marked for regeneration."
									  : L"Nothing would be written.",
								  WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX, CRect (), this, IdFoot);
				m_footCtl.SetFont (&m_font);

				m_list.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | LVS_REPORT
								   | LVS_SINGLESEL | LVS_NOSORTHEADER | LVS_SHOWSELALWAYS,
							   CRect (), this, IdList);
				m_list.SetExtendedStyle (LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_INFOTIP);
				m_list.SetFont (&m_font);
				// Taller rows, for air: a list takes its row height from its images.
				m_rowImages.Create (1, m_u * 8 / 5, ILC_COLOR, 1, 0);
				m_list.SetImageList (&m_rowImages, LVSIL_SMALL);

				m_list.InsertColumn (0, L"Operation", LVCFMT_LEFT, m_u * 9);
				m_list.InsertColumn (1, L"Parameter", LVCFMT_LEFT, m_u * 11);
				m_list.InsertColumn (2, L"Now", LVCFMT_LEFT, m_u * 14);
				m_list.InsertColumn (3, L"Will be", LVCFMT_LEFT, m_u * 14);

				for (size_t i = 0; i < m_lines.size (); ++i)
					{
					const Preview::Line &l = m_lines[i];
					const int at = m_list.InsertItem (static_cast<int> (i),
													  l.kind == Preview::Line::Change ? L"" : l.text.c_str ());
					if (l.kind == Preview::Line::Change)
						{
						m_list.SetItemText (at, 1, l.text.c_str ());
						m_list.SetItemText (at, 2, Show1 (l.from).c_str ());
						m_list.SetItemText (at, 3, Show1 (l.to).c_str ());
						}
					else
						m_list.SetItemText (at, 1, l.detail.c_str ());
					}

				const std::wstring apply = m_changes > 0
					? L"Apply " + std::to_wstring (m_changes) + (m_changes == 1 ? L" change" : L" changes")
					: L"Apply";
				m_apply.Create (apply.c_str (), WS_CHILD | WS_TABSTOP | BS_DEFPUSHBUTTON
								| (m_changes > 0 ? WS_VISIBLE : 0), CRect (), this, IDOK);
				m_apply.SetFont (&m_bold);
				m_cancel.Create (m_changes > 0 ? L"Cancel" : L"Close",
								 WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, CRect (), this, IDCANCEL);
				m_cancel.SetFont (&m_font);

				// A comfortable size on this screen, centred over Mastercam.
				CRect work;
				SystemParametersInfo (SPI_GETWORKAREA, 0, &work, 0);
				const int w = (std::min) (m_u * 64, work.Width () * 9 / 10);
				const int h = (std::min) (m_u * 40, work.Height () * 9 / 10);
				SetWindowPos (nullptr, work.left + (work.Width () - w) / 2,
							  work.top + (work.Height () - h) / 2, w, h, SWP_NOZORDER);
				Layout ();

				(m_changes > 0 ? m_apply : m_cancel).SetFocus ();
				return FALSE;
				}

			void Layout ()
				{
				if (m_list.GetSafeHwnd () == nullptr)
					return;
				CRect rc;
				GetClientRect (&rc);
				const int u = m_u;
				const int pad = u, gap = u / 2, btnW = m_btnW, btnH = u * 2;
				const int titleH = m_bigH + u / 4, summaryH = u * 3 / 2, footH = u * 3 / 2;
				int y = pad;
				m_titleCtl.MoveWindow (pad, y, rc.Width () - 2 * pad, titleH);
				y += titleH;
				m_summaryCtl.MoveWindow (pad, y, rc.Width () - 2 * pad, summaryH);
				y += summaryH + gap;
				const int bottom = rc.bottom - pad - btnH - gap - footH - gap;
				m_list.MoveWindow (pad, y, rc.Width () - 2 * pad, (std::max) (60, bottom - y));
				m_footCtl.MoveWindow (pad, bottom + gap, rc.Width () - 2 * pad, footH);
				m_cancel.MoveWindow (rc.right - pad - btnW, rc.bottom - pad - btnH, btnW, btnH);
				m_apply.MoveWindow (rc.right - pad - 2 * btnW - gap, rc.bottom - pad - btnH, btnW, btnH);

				// The two value columns share whatever width the first two leave.
				CRect lr;
				m_list.GetClientRect (&lr);
				const int fixed = m_list.GetColumnWidth (0) + m_list.GetColumnWidth (1);
				const int each = (std::max) (m_u * 6, (lr.Width () - fixed - 2) / 2);
				m_list.SetColumnWidth (2, each);
				m_list.SetColumnWidth (3, each);
				}

			afx_msg void OnSize (UINT type, int cx, int cy)
				{
				CDialog::OnSize (type, cx, cy);
				Layout ();
				}

			afx_msg void OnGetMinMaxInfo (MINMAXINFO *mmi)
				{
				mmi->ptMinTrackSize.x = m_u * 40;
				mmi->ptMinTrackSize.y = m_u * 22;
				}

			afx_msg HBRUSH OnCtlColor (CDC *dc, CWnd *wnd, UINT ctl)
				{
				HBRUSH br = CDialog::OnCtlColor (dc, wnd, ctl);
				if (wnd != nullptr && (wnd->GetDlgCtrlID () == IdSummary || wnd->GetDlgCtrlID () == IdFoot))
					dc->SetTextColor (kMuted);
				if (wnd != nullptr && wnd->GetDlgCtrlID () == IdTitle)
					dc->SetTextColor (kInk);
				return br;
				}

			/// Section, operation, refused and note rows are drawn whole - their
			/// text runs across the columns. Change rows are drawn by the list,
			/// with a colour and font per column.
			afx_msg void OnCustomDraw (NMHDR *hdr, LRESULT *result)
				{
				NMLVCUSTOMDRAW *cd = reinterpret_cast<NMLVCUSTOMDRAW *> (hdr);
				*result = CDRF_DODEFAULT;
				switch (cd->nmcd.dwDrawStage)
					{
					case CDDS_PREPAINT:
						*result = CDRF_NOTIFYITEMDRAW;
						return;

					case CDDS_ITEMPREPAINT:
						{
						const size_t i = static_cast<size_t> (cd->nmcd.dwItemSpec);
						if (i >= m_lines.size ())
							return;
						const Preview::Line &l = m_lines[i];
						if (l.kind == Preview::Line::Change)
							{
							*result = CDRF_NOTIFYSUBITEMDRAW;
							return;
							}
						DrawWhole (cd->nmcd.hdc, static_cast<int> (i), l);
						*result = CDRF_SKIPDEFAULT;
						return;
						}

					case CDDS_ITEMPREPAINT | CDDS_SUBITEM:
						{
						const size_t i = static_cast<size_t> (cd->nmcd.dwItemSpec);
						const bool selected = m_list.GetItemState (static_cast<int> (i), LVIS_SELECTED) != 0;
						switch (cd->iSubItem)
							{
							case 2:
								cd->clrText = kOldInk;
								cd->clrTextBk = selected ? RGB (0xF6, 0xD4, 0xD0) : kOldBg;
								SelectObject (cd->nmcd.hdc, m_strike.GetSafeHandle ());
								break;
							case 3:
								cd->clrText = kNewInk;
								cd->clrTextBk = selected ? RGB (0xC9, 0xEB, 0xD7) : kNewBg;
								SelectObject (cd->nmcd.hdc, m_bold.GetSafeHandle ());
								break;
							default:
								cd->clrText = kInk;
								cd->clrTextBk = selected ? RGB (0xDC, 0xE6, 0xF5) : kWhite;
								SelectObject (cd->nmcd.hdc, m_font.GetSafeHandle ());
								break;
							}
						// The list paints a selected row in the system colour unless
						// told otherwise; keep the meaning colours.
						cd->nmcd.uItemState &= ~(CDIS_SELECTED | CDIS_FOCUS);
						*result = CDRF_NEWFONT;
						return;
						}
					}
				}

			void DrawWhole (HDC hdc, int item, const Preview::Line &l)
				{
				CRect rc;
				m_list.GetItemRect (item, &rc, LVIR_BOUNDS);
				CDC *dc = CDC::FromHandle (hdc);

				COLORREF bg = kWhite, ink = kInk, ink2 = kMuted;
				CFont *font = &m_font, *font2 = &m_font;
				switch (l.kind)
					{
					case Preview::Line::Section:
						bg = kSectionBg; font = &m_bold;
						break;
					case Preview::Line::Op:
						bg = kOpBg; font = &m_bold;
						break;
					case Preview::Line::Refused:
						bg = kRefusedBg; ink = kRefusedInk; ink2 = kRefusedInk; font = &m_bold;
						break;
					default:
						ink = kMuted;
						break;
					}
				dc->FillSolidRect (&rc, bg);
				if (l.kind == Preview::Line::Section || l.kind == Preview::Line::Op)
					{
					CRect line = rc;
					line.top = line.bottom - 1;
					dc->FillSolidRect (&line, RGB (0xD5, 0xDC, 0xE6));
					}
				dc->SetBkMode (TRANSPARENT);

				// The label, then its detail right after it - both across the full
				// width, so neither is cut to a column. The detail starts no
				// earlier than the Parameter column, so details line up.
				CRect a = rc;
				a.left += 8;
				a.right -= 8;
				CFont *old = dc->SelectObject (font);
				dc->SetTextColor (ink);
				CRect measured = a;
				dc->DrawText (l.text.c_str (), static_cast<int> (l.text.size ()), &measured,
							  DT_SINGLELINE | DT_NOPREFIX | DT_CALCRECT);
				dc->DrawText (l.text.c_str (), static_cast<int> (l.text.size ()), &a,
							  DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
				if (l.kind != Preview::Line::Section && !l.detail.empty ())
					{
					CRect b = a;
					b.left = (std::max) (static_cast<int> (measured.right) + 24,
										 static_cast<int> (rc.left) + m_list.GetColumnWidth (0) + 4);
					dc->SelectObject (font2);
					dc->SetTextColor (ink2);
					if (b.left < b.right)
						dc->DrawText (l.detail.c_str (), static_cast<int> (l.detail.size ()), &b,
									  DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX);
					}
				dc->SelectObject (old);
				}

			DECLARE_MESSAGE_MAP ()

		private:
			std::wstring m_title, m_summary;
			const std::vector<Preview::Line> &m_lines;
			int m_changes = 0;
			int m_u = 16;				//!< one line of text, in pixels - every size derives from it
			int m_bigH = 24;
			int m_btnW = 150;
			CWnd *m_parent = nullptr;
			std::vector<WORD> m_tpl;

			CFont m_font, m_bold, m_strike, m_big;
			CStatic m_titleCtl, m_summaryCtl, m_footCtl;
			CListCtrl m_list;
			CImageList m_rowImages;
			CButton m_apply, m_cancel;
		};

	BEGIN_MESSAGE_MAP (PreviewDlg, CDialog)
		ON_WM_SIZE ()
		ON_WM_GETMINMAXINFO ()
		ON_WM_CTLCOLOR ()
		ON_NOTIFY (NM_CUSTOMDRAW, IdList, &PreviewDlg::OnCustomDraw)
	END_MESSAGE_MAP ()
	}

namespace Preview
	{
	bool Show (const std::wstring &title, const std::wstring &summary,
			   const std::vector<Line> &lines, int changes)
		{
		PreviewDlg dlg (title, summary, lines, changes,
						CWnd::FromHandle (get_MainFrame ()->GetSafeHwnd ()));
		return dlg.Run () == IDOK && changes > 0;
		}
	}
