// The window itself - no Mastercam SDK (Ui.h), so it can be drawn outside
// Mastercam for the manual's pictures (tests\dialog_shots.cpp).
#include "Ui.h"
#include "Preview.h"

#include <algorithm>

namespace
	{
	// ---- Colours: light, quiet, one meaning each.
	const COLORREF kInk        = RGB (0x1F, 0x29, 0x37);
	const COLORREF kMuted      = RGB (0x5B, 0x65, 0x73);
	const COLORREF kLeftOut    = RGB (0xA0, 0xA6, 0xAE);
	const COLORREF kSectionBg  = RGB (0xE6, 0xEB, 0xF2);
	const COLORREF kOpBg       = RGB (0xF4, 0xF6, 0xFA);
	const COLORREF kOldInk     = RGB (0xB4, 0x23, 0x18);
	const COLORREF kOldBg      = RGB (0xFD, 0xEE, 0xEC);
	const COLORREF kNewInk     = RGB (0x06, 0x76, 0x47);
	const COLORREF kNewBg      = RGB (0xE8, 0xF6, 0xEE);
	const COLORREF kRefusedInk = RGB (0x8A, 0x4B, 0x00);
	const COLORREF kRefusedBg  = RGB (0xFF, 0xF5, 0xE1);
	const COLORREF kWhite      = RGB (0xFF, 0xFF, 0xFF);
	const COLORREF kAccent     = RGB (0x25, 0x63, 0xEB);	//!< a ticked box

	enum { IdList = 1001, IdTitle = 1002, IdSummary = 1003, IdFoot = 1004, IdImpact = 1005 };

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

	COLORREF ToneInk (int tone)
		{
		return tone < 0 ? kNewInk : tone > 0 ? kOldInk : kInk;
		}

	class PreviewDlg : public CDialog
		{
		public:
			PreviewDlg (const std::wstring &title, const std::wstring &summary,
						std::vector<Preview::Line> &lines, const Preview::Options &options, CWnd *parent)
				: m_title (title), m_summary (summary), m_lines (lines), m_opt (options),
				  m_parent (parent)
				{
				for (const Preview::Line &l : m_lines)
					if (l.kind == Preview::Line::Change && l.box != Preview::Line::NoBox)
						++m_offered;
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
				SetWindowText (m_opt.caption.c_str ());

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
				LOGFONT mid = lf;
				mid.lfHeight = lf.lfHeight * 6 / 5;
				mid.lfWeight = FW_SEMIBOLD;
				m_mid.CreateFontIndirect (&mid);
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
				dc.SelectObject (&m_mid);
				dc.GetTextMetrics (&tm);
				m_midH = tm.tmHeight;
				dc.SelectObject (&m_bold);
				const std::wstring longest = m_opt.verb + L" 9999 changes";
				m_btnW = dc.GetTextExtent (longest.c_str (), static_cast<int> (longest.size ())).cx
						 + 2 * m_u;
				dc.SelectObject (old);
				}
				m_box = (std::max) (11, m_u * 4 / 5);

				m_titleCtl.Create (m_title.c_str (), WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,
								   CRect (), this, IdTitle);
				m_titleCtl.SetFont (&m_big);
				m_summaryCtl.Create (m_summary.c_str (), WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,
									 CRect (), this, IdSummary);
				m_summaryCtl.SetFont (&m_font);
				m_impactCtl.Create (L"", WS_CHILD | (m_opt.impact ? WS_VISIBLE : 0) | SS_LEFT | SS_NOPREFIX
									| SS_ENDELLIPSIS, CRect (), this, IdImpact);
				m_impactCtl.SetFont (&m_mid);
				m_footCtl.Create (m_offered > 0 ? m_opt.foot.c_str () : m_opt.nothing.c_str (),
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

				m_apply.Create (m_opt.verb.c_str (), WS_CHILD | WS_TABSTOP | BS_DEFPUSHBUTTON
								| (m_offered > 0 ? WS_VISIBLE : 0), CRect (), this, IDOK);
				m_apply.SetFont (&m_bold);
				m_cancel.Create (m_offered > 0 ? L"Cancel" : L"Close",
								 WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, CRect (), this, IDCANCEL);
				m_cancel.SetFont (&m_font);

				// A comfortable size on this screen, centred over Mastercam.
				CRect work;
				SystemParametersInfo (SPI_GETWORKAREA, 0, &work, 0);
				const int w = (std::min) (m_u * 68, work.Width () * 9 / 10);
				const int h = (std::min) (m_u * 42, work.Height () * 9 / 10);
				SetWindowPos (nullptr, work.left + (work.Width () - w) / 2,
							  work.top + (work.Height () - h) / 2, w, h, SWP_NOZORDER);
				Preview::Sync (m_lines);
				Refresh ();
				Layout ();

				(m_offered > 0 ? m_apply : m_cancel).SetFocus ();
				return FALSE;
				}

			/// After any tick: the impact, the button's count, the list redrawn.
			void Refresh ()
				{
				if (m_opt.impact)
					{
					m_impactTone = 0;
					const std::wstring text = m_opt.impact (m_lines, m_impactTone);
					m_impactCtl.SetWindowText (text.c_str ());
					m_impactCtl.Invalidate ();
					}
				const int n = Preview::CountTicked (m_lines);
				const std::wstring apply = m_opt.verb + L" " + std::to_wstring (n)
										   + (n == 1 ? L" change" : L" changes");
				m_apply.SetWindowText (apply.c_str ());
				m_apply.EnableWindow (n > 0);
				if (m_list.GetSafeHwnd () != nullptr)
					m_list.Invalidate (FALSE);
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
				const int impactH = m_opt.impact ? m_midH + u / 2 : 0;
				int y = pad;
				m_titleCtl.MoveWindow (pad, y, rc.Width () - 2 * pad, titleH);
				y += titleH;
				m_summaryCtl.MoveWindow (pad, y, rc.Width () - 2 * pad, summaryH);
				y += summaryH;
				m_impactCtl.MoveWindow (pad, y, rc.Width () - 2 * pad, impactH);
				y += impactH + gap;
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
				mmi->ptMinTrackSize.y = m_u * 24;
				}

			afx_msg HBRUSH OnCtlColor (CDC *dc, CWnd *wnd, UINT ctl)
				{
				HBRUSH br = CDialog::OnCtlColor (dc, wnd, ctl);
				const int id = wnd != nullptr ? wnd->GetDlgCtrlID () : 0;
				if (id == IdSummary || id == IdFoot)
					dc->SetTextColor (kMuted);
				if (id == IdTitle)
					dc->SetTextColor (kInk);
				if (id == IdImpact)
					dc->SetTextColor (ToneInk (m_impactTone));
				return br;
				}

			// ---- Tick boxes ------------------------------------------------

			/// Where line `item`'s box is drawn, or an empty rect for none. An
			/// operation's box sits at the left; its changes' boxes one step in.
			CRect BoxRect (int item)
				{
				const size_t i = static_cast<size_t> (item);
				if (i >= m_lines.size () || m_lines[i].box == Preview::Line::NoBox)
					return CRect ();
				CRect rc;
				m_list.GetItemRect (item, &rc, LVIR_BOUNDS);
				const int left = rc.left + 8 + (m_lines[i].kind == Preview::Line::Change ? m_u : 0);
				const int top = rc.top + (rc.Height () - m_box) / 2;
				return CRect (left, top, left + m_box, top + m_box);
				}

			/// Drawn here rather than by the system: the system's "mixed" box is
			/// a check in grey, which at this size reads as ticked - and an
			/// operation partly left out must not look wholly applied.
			void DrawBox (CDC *dc, const CRect &r, Preview::Line::Box box)
				{
				if (r.IsRectEmpty ())
					return;
				const bool on = box == Preview::Line::Ticked || box == Preview::Line::Mixed;
				dc->FillSolidRect (&r, on ? kAccent : kMuted);
				CRect in = r;
				in.DeflateRect (1, 1);
				if (!on)
					{
					dc->FillSolidRect (&in, kWhite);
					return;
					}
				const int w = r.Width (), h = r.Height ();
				const int th = (std::max) (2, w / 6);
				if (box == Preview::Line::Mixed)
					{
					// A bar: some of it.
					CRect bar (r.left + w / 4, r.top + (h - th) / 2, r.right - w / 4, r.top + (h - th) / 2 + th);
					dc->FillSolidRect (&bar, kWhite);
					return;
					}
				CPen pen (PS_SOLID, th, kWhite);
				CPen *old = dc->SelectObject (&pen);
				const POINT tick[3] = { { r.left + w * 22 / 100, r.top + h * 52 / 100 },
										{ r.left + w * 42 / 100, r.top + h * 72 / 100 },
										{ r.left + w * 78 / 100, r.top + h * 30 / 100 } };
				dc->Polyline (tick, 3);
				dc->SelectObject (old);
				}

			void ToggleItem (int item)
				{
				if (item < 0 || static_cast<size_t> (item) >= m_lines.size ()
					|| m_lines[static_cast<size_t> (item)].box == Preview::Line::NoBox)
					return;
				Preview::Toggle (m_lines, static_cast<size_t> (item));
				Refresh ();
				}

			/// Whether x on line `item` is on its box. On a change line the whole
			/// first column counts - it holds nothing else, and a box is a small
			/// target.
			bool OnBox (int item, int x)
				{
				const CRect box = BoxRect (item);
				if (box.IsRectEmpty ())
					return false;
				CRect row;
				m_list.GetItemRect (item, &row, LVIR_BOUNDS);
				const bool change = m_lines[static_cast<size_t> (item)].kind == Preview::Line::Change;
				const int right = change ? row.left + m_list.GetColumnWidth (0) : box.right + m_u / 2;
				return x >= row.left && x < right;
				}

			afx_msg void OnClick (NMHDR *hdr, LRESULT *result)
				{
				*result = 0;
				const NMITEMACTIVATE *a = reinterpret_cast<const NMITEMACTIVATE *> (hdr);
				if (OnBox (a->iItem, a->ptAction.x))
					ToggleItem (a->iItem);
				}

			/// A double-click anywhere else on a line ticks it too. On the box it
			/// would tick twice - the first click of the pair already did.
			afx_msg void OnDblClick (NMHDR *hdr, LRESULT *result)
				{
				*result = 0;
				const NMITEMACTIVATE *a = reinterpret_cast<const NMITEMACTIVATE *> (hdr);
				if (a->iItem >= 0 && !OnBox (a->iItem, a->ptAction.x))
					ToggleItem (a->iItem);
				}

			afx_msg void OnKeyDown (NMHDR *hdr, LRESULT *result)
				{
				*result = 0;
				const NMLVKEYDOWN *k = reinterpret_cast<const NMLVKEYDOWN *> (hdr);
				if (k->wVKey == VK_SPACE)
					ToggleItem (m_list.GetNextItem (-1, LVNI_SELECTED));
				}

			/// Section, operation, refused and note rows are drawn whole - their
			/// text runs across the columns. Change rows are drawn by the list,
			/// with a colour and font per column, and their box drawn here.
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
						if (i >= m_lines.size ())
							return;
						const bool selected = m_list.GetItemState (static_cast<int> (i), LVIS_SELECTED) != 0;
						const bool off = m_lines[i].box == Preview::Line::Unticked;
						if (cd->iSubItem == 0)
							{
							// The first column holds only the box.
							CRect row;
							m_list.GetItemRect (static_cast<int> (i), &row, LVIR_BOUNDS);
							row.right = row.left + m_list.GetColumnWidth (0);
							CDC *dc = CDC::FromHandle (cd->nmcd.hdc);
							dc->FillSolidRect (&row, selected ? RGB (0xDC, 0xE6, 0xF5) : kWhite);
							DrawBox (dc, BoxRect (static_cast<int> (i)), m_lines[i].box);
							*result = CDRF_SKIPDEFAULT;
							return;
							}
						switch (cd->iSubItem)
							{
							case 2:
								cd->clrText = off ? kLeftOut : kOldInk;
								cd->clrTextBk = off ? kWhite : selected ? RGB (0xF6, 0xD4, 0xD0) : kOldBg;
								SelectObject (cd->nmcd.hdc, m_strike.GetSafeHandle ());
								break;
							case 3:
								cd->clrText = off ? kLeftOut : kNewInk;
								cd->clrTextBk = off ? kWhite : selected ? RGB (0xC9, 0xEB, 0xD7) : kNewBg;
								SelectObject (cd->nmcd.hdc, (off ? m_font : m_bold).GetSafeHandle ());
								break;
							default:
								cd->clrText = off ? kLeftOut : kInk;
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
						if (l.box == Preview::Line::Unticked)
							ink = ink2 = kLeftOut;
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

				CRect a = rc;
				a.left += 8;
				a.right -= 8;
				const CRect box = BoxRect (item);
				if (!box.IsRectEmpty ())
					{
					DrawBox (dc, box, l.box);
					a.left = box.right + m_u / 2;
					}

				// The impact at the right, so the eye can run down the column of
				// them; the label and its detail have what is left.
				CFont *old = dc->SelectObject (&m_bold);
				if (!l.impact.empty ())
					{
					CRect m = a;
					dc->DrawText (l.impact.c_str (), static_cast<int> (l.impact.size ()), &m,
								  DT_SINGLELINE | DT_NOPREFIX | DT_CALCRECT);
					CRect at = a;
					at.left = (std::max) (a.left + a.Width () / 3, a.right - m.Width ());
					dc->SetTextColor (l.box == Preview::Line::Unticked ? kLeftOut : ToneInk (l.tone));
					dc->DrawText (l.impact.c_str (), static_cast<int> (l.impact.size ()), &at,
								  DT_SINGLELINE | DT_VCENTER | DT_RIGHT | DT_END_ELLIPSIS | DT_NOPREFIX);
					a.right = at.left - 24;
					}

				// The label, then its detail right after it - both across the full
				// width, so neither is cut to a column. The detail starts no
				// earlier than the Parameter column, so details line up.
				dc->SelectObject (font);
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
			std::vector<Preview::Line> &m_lines;
			const Preview::Options &m_opt;
			int m_offered = 0;			//!< change lines with a box - 0 = nothing to apply
			int m_impactTone = 0;
			int m_u = 16;				//!< one line of text, in pixels - every size derives from it
			int m_bigH = 24;
			int m_midH = 20;
			int m_btnW = 150;
			int m_box = 13;
			CWnd *m_parent = nullptr;
			std::vector<WORD> m_tpl;

			CFont m_font, m_bold, m_strike, m_big, m_mid;
			CStatic m_titleCtl, m_summaryCtl, m_impactCtl, m_footCtl;
			CListCtrl m_list;
			CImageList m_rowImages;
			CButton m_apply, m_cancel;
		};

	BEGIN_MESSAGE_MAP (PreviewDlg, CDialog)
		ON_WM_SIZE ()
		ON_WM_GETMINMAXINFO ()
		ON_WM_CTLCOLOR ()
		ON_NOTIFY (NM_CUSTOMDRAW, IdList, &PreviewDlg::OnCustomDraw)
		ON_NOTIFY (NM_CLICK, IdList, &PreviewDlg::OnClick)
		ON_NOTIFY (NM_DBLCLK, IdList, &PreviewDlg::OnDblClick)
		ON_NOTIFY (LVN_KEYDOWN, IdList, &PreviewDlg::OnKeyDown)
	END_MESSAGE_MAP ()
	}

namespace Preview
	{
	bool Show (const std::wstring &title, const std::wstring &summary,
			   std::vector<Line> &lines, const Options &options)
		{
		PreviewDlg dlg (title, summary, lines, options, Ui::Host ());
		return dlg.Run () == IDOK && CountTicked (lines) > 0;
		}
	}
