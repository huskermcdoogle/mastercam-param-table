// The batch dump's windows - no Mastercam SDK (Ui.h), as the dump window, so they
// build and show outside Mastercam too.
#include "Ui.h"
#include "BatchDialog.h"
#include "Shop.h"

#include <algorithm>
#include <ctime>
#include <filesystem>

namespace
	{
	const COLORREF kInk   = RGB (0x1F, 0x29, 0x37);
	const COLORREF kMuted = RGB (0x5B, 0x65, 0x73);
	const COLORREF kGood  = RGB (0x06, 0x76, 0x47);
	const COLORREF kBad   = RGB (0xB4, 0x23, 0x18);
	const COLORREF kWarn  = RGB (0x8A, 0x4B, 0x00);		// the open part is closed: read this first

	/// How many parts the window counts before it stops looking and says "more than":
	/// a folder pointed at by mistake (a whole drive) must not hang the window.
	const size_t kCountLimit = 2000;

	enum
		{
		IdTitle = 2101, IdWarn, IdPartsHead, IdFolder, IdBrowse, IdSub, IdCount, IdOutHead, IdBeside, IdOne,
		IdOut, IdOutBrowse, IdOptHead, IdPictures, IdMacros, IdStockSim, IdShop
		};
	const UINT_PTR kCountTimer = 1;

	/// The dialog fonts: the system's message font, a bold one and a big one for the title.
	struct Fonts
		{
		CFont font, bold, big;
		int u = 16, bigH = 24;

		void Make (CWnd &w)
			{
			NONCLIENTMETRICS ncm = { sizeof (ncm) };
			SystemParametersInfo (SPI_GETNONCLIENTMETRICS, sizeof (ncm), &ncm, 0);
			LOGFONT lf = ncm.lfMessageFont;
			font.CreateFontIndirect (&lf);
			LOGFONT b = lf;
			b.lfWeight = FW_SEMIBOLD;
			bold.CreateFontIndirect (&b);
			LOGFONT g = lf;
			g.lfHeight = lf.lfHeight * 3 / 2;
			g.lfWeight = FW_SEMIBOLD;
			big.CreateFontIndirect (&g);
			w.SetFont (&font);
			CClientDC dc (&w);
			TEXTMETRIC tm;
			CFont *old = dc.SelectObject (&font);
			dc.GetTextMetrics (&tm);
			u = (std::max) (12, static_cast<int> (tm.tmHeight));
			dc.SelectObject (&big);
			dc.GetTextMetrics (&tm);
			bigH = tm.tmHeight;
			dc.SelectObject (old);
			}
		};

	class Dlg : public CDialog
		{
		public:
			Dlg (const std::wstring &openPart, const std::wstring &ext, Settings::Batch &settings, CWnd *parent)
				: m_openPart (openPart), m_ext (ext), m_settings (settings), m_parent (parent)
				{
				}

			INT_PTR Run ()
				{
				m_tpl.assign (64, 0);
				DLGTEMPLATE *t = reinterpret_cast<DLGTEMPLATE *> (m_tpl.data ());
				t->style = WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME | DS_CENTER;
				t->cx = 300;
				t->cy = 200;
				InitModalIndirect (t, m_parent);
				return DoModal ();
				}

		protected:
			BOOL OnInitDialog () override
				{
				CDialog::OnInitDialog ();
				SetWindowText (L"Parameter Table Tool - dump a folder of parts");
				m_f.Make (*this);

				auto label = [&] (CStatic &c, const std::wstring &text, int id, CFont &f, DWORD extra = SS_LEFTNOWORDWRAP | SS_ENDELLIPSIS)
					{
					c.Create (text.c_str (), WS_CHILD | WS_VISIBLE | SS_NOPREFIX | extra, CRect (), this, id);
					c.SetFont (&f);
					};
				auto button = [&] (CButton &b, const wchar_t *text, int id, DWORD style)
					{
					b.Create (text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | style, CRect (), this, id);
					b.SetFont (&m_f.font);
					};
				auto edit = [&] (CEdit &e, const std::wstring &text, int id)
					{
					e.Create (WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL, CRect (), this, id);
					e.SetFont (&m_f.font);
					e.SetWindowText (text.c_str ());
					};

				const std::filesystem::path open (m_openPart);
				label (m_title, L"Dump a folder of parts", IdTitle, m_f.big);
				// What happens to the part on screen, first: it is closed.
				label (m_warn, m_openPart.empty ()
								   ? L"Mastercam has one part open at a time. Each part of the folder is opened in turn - "
									 L"not regenerated, never saved - dumped like \"Lathe params - dump to Excel\" "
									 L"(every op, no window), and closed. An empty part is left open at the end."
								   : L"Mastercam has one part open at a time, so " + open.filename ().wstring ()
									 + L" is closed while this runs. Each part of the folder is opened in turn - "
									   L"not regenerated, never saved - dumped like \"Lathe params - dump to Excel\" "
									   L"(every op, no window), and closed. " + open.filename ().wstring ()
									 + L" is opened again at the end.",
					   IdWarn, m_f.font, 0);

				label (m_partsHead, L"Parts", IdPartsHead, m_f.bold);
				edit (m_folder, !m_settings.folder.empty () ? m_settings.folder
							   : !m_openPart.empty () ? open.parent_path ().wstring () : std::wstring (), IdFolder);
				button (m_browse, L"Browse...", IdBrowse, BS_PUSHBUTTON);
				button (m_sub, L"Include subfolders", IdSub, BS_AUTOCHECKBOX);
				m_sub.SetCheck (m_settings.subfolders ? BST_CHECKED : BST_UNCHECKED);
				label (m_count, L"", IdCount, m_f.bold);

				label (m_outHead, L"Save the workbooks", IdOutHead, m_f.bold);
				button (m_beside, L"Beside each part, as a normal dump", IdBeside, BS_AUTORADIOBUTTON | WS_GROUP);
				button (m_one, L"All in one folder:", IdOne, BS_AUTORADIOBUTTON);
				edit (m_out, m_settings.outFolder, IdOut);
				button (m_outBrowse, L"Browse...", IdOutBrowse, BS_PUSHBUTTON);
				(m_settings.outFolder.empty () ? m_beside : m_one).SetCheck (BST_CHECKED);

				label (m_optHead, L"Options", IdOptHead, m_f.bold);
				button (m_pics, L"Tool pictures on the Tools sheet (lathe tools)", IdPictures, BS_AUTOCHECKBOX | WS_GROUP);
				m_pics.SetCheck (m_settings.pictures ? BST_CHECKED : BST_UNCHECKED);
				button (m_macros, L"Include macros (.xlsm): the Parameter Table ribbon tab in each workbook", IdMacros,
						BS_AUTOCHECKBOX);
				m_macros.SetCheck (m_settings.macros ? BST_CHECKED : BST_UNCHECKED);
				button (m_sim, L"Simulate the stock (slow - off: Mastercam's own stock boundaries)", IdStockSim,
						BS_AUTOCHECKBOX);
				m_sim.SetCheck (m_settings.stockSim ? BST_CHECKED : BST_UNCHECKED);

				label (m_shop, L"", IdShop, m_f.font);

				button (m_ok, L"Dump", IDOK, BS_DEFPUSHBUTTON);
				m_ok.SetFont (&m_f.bold);
				button (m_cancel, L"Cancel", IDCANCEL, BS_PUSHBUTTON);

				CRect work;
				SystemParametersInfo (SPI_GETWORKAREA, 0, &work, 0);
				const int w = (std::max) (m_f.u * 40, Needed ());
				const int h = Layout (w, true);
				CRect wr (0, 0, w, h);
				CalcWindowRect (&wr);
				SetWindowPos (nullptr, work.left + (work.Width () - wr.Width ()) / 2,
							  work.top + (work.Height () - wr.Height ()) / 2, wr.Width (), wr.Height (), SWP_NOZORDER);
				Layout (w, false);
				// The folders from their start: filled in while the boxes had no width,
				// they were left scrolled to the end.
				m_folder.SetSel (0, 0);
				m_out.SetSel (0, 0);
				m_ready = true;
				Count ();
				m_ok.SetFocus ();
				return FALSE;
				}

			/// Place everything for a client width; returns the height used.
			int Layout (int w, bool measureOnly)
				{
				const int u = m_f.u, pad = u, gap = u / 2, inner = w - 2 * pad;
				const int lineH = u * 3 / 2, editH = u * 3 / 2, btnH = u * 2;
				int y = pad;
				auto place = [&] (CWnd &c, int x, int yy, int cw, int ch)
					{
					if (!measureOnly)
						c.MoveWindow (x, yy, cw, ch);
					};
				place (m_title, pad, y, inner, m_f.bigH + 2);
				y += m_f.bigH + 2 + gap;
				const int warnH = TextH (m_warn, inner);
				place (m_warn, pad, y, inner, warnH);
				y += warnH + gap * 2;

				const int bw = ButtonW (m_browse);
				place (m_partsHead, pad, y, inner, lineH);
				y += lineH;
				place (m_folder, pad + u, y, inner - u - bw - gap, editH);
				place (m_browse, w - pad - bw, y, bw, editH);
				y += editH + gap / 2;
				const int subW = CheckW (m_sub);
				place (m_sub, pad + u, y, subW, lineH);
				place (m_count, pad + u + subW + u, y, (std::max) (0, w - pad - (pad + u + subW + u)), lineH);
				y += lineH + gap;

				place (m_outHead, pad, y, inner, lineH);
				y += lineH;
				place (m_beside, pad + u, y, inner - u, lineH);
				y += lineH;
				const int oneW = CheckW (m_one);
				place (m_one, pad + u, y, oneW, editH);
				place (m_out, pad + u + oneW + gap, y, (std::max) (u * 8, w - pad - bw - gap - (pad + u + oneW + gap)), editH);
				place (m_outBrowse, w - pad - bw, y, bw, editH);
				y += editH + gap;

				place (m_optHead, pad, y, inner, lineH);
				y += lineH;
				for (CButton *b : { &m_pics, &m_macros, &m_sim })
					{
					place (*b, pad + u, y, inner - u, lineH);
					y += lineH;
					}
				y += gap;
				place (m_shop, pad, y, inner, lineH);
				y += lineH + gap;

				const int okW = u * 11, cW = u * 7;
				place (m_ok, w - pad - okW - gap - cW, y, okW, btnH);
				place (m_cancel, w - pad - cW, y, cW, btnH);
				y += btnH + pad;
				return y;
				}

			int TextW (const CString &text)
				{
				CClientDC dc (this);
				CFont *old = dc.SelectObject (&m_f.font);
				const int cx = dc.GetTextExtent (text).cx;
				dc.SelectObject (old);
				return cx;
				}

			/// The height wrapped text takes at a width.
			int TextH (CWnd &c, int width)
				{
				CString text;
				c.GetWindowText (text);
				CClientDC dc (this);
				CFont *old = dc.SelectObject (&m_f.font);
				CRect r (0, 0, width, 0);
				dc.DrawText (text, &r, DT_CALCRECT | DT_WORDBREAK | DT_NOPREFIX);
				dc.SelectObject (old);
				return r.Height () + 2;
				}

			int ButtonW (CWnd &c)
				{
				CString text;
				c.GetWindowText (text);
				return TextW (text) + m_f.u * 2;
				}

			/// A tick box or radio button sized to its words.
			int CheckW (CWnd &c)
				{
				CString text;
				c.GetWindowText (text);
				return TextW (text) + GetSystemMetrics (SM_CXMENUCHECK) + m_f.u;
				}

			/// Wide enough for the longest option at this font.
			int Needed ()
				{
				int widest = 0;
				for (CButton *b : { &m_pics, &m_macros, &m_sim, &m_beside })
					widest = (std::max) (widest, CheckW (*b));
				return widest + 3 * m_f.u;
				}

			std::wstring FolderText (CEdit &e)
				{
				CString t;
				e.GetWindowText (t);
				std::wstring s = t.GetString ();
				while (!s.empty () && (s.back () == L' ' || s.back () == L'\t'))
					s.pop_back ();
				while (!s.empty () && (s.front () == L' ' || s.front () == L'\t'))
					s.erase (s.begin ());
				return s;
				}

			/// How many parts the folder holds - the button and the line under the
			/// folder say it - and where the shop workbook will go.
			void Count ()
				{
				if (!m_ready)
					return;
				const std::wstring folder = FolderText (m_folder);
				std::error_code ec;
				m_found = 0;
				m_more = false;
				if (folder.empty () || !std::filesystem::is_directory (folder, ec))
					{
					m_countOk = false;
					m_count.SetWindowText (folder.empty () ? L"pick the folder of parts" : L"that folder does not exist");
					}
				else
					{
					CWaitCursor wait;
					m_found = Shop::FindParts (folder, m_sub.GetCheck () == BST_CHECKED, m_ext, kCountLimit, m_more).size ();
					m_countOk = m_found > 0;
					m_count.SetWindowText (m_found == 0 ? (L"no " + m_ext + L" parts there").c_str ()
										   : m_more ? (L"more than " + std::to_wstring (kCountLimit) + L" parts - is this the right folder?").c_str ()
										   : (std::to_wstring (m_found) + (m_found == 1 ? L" part found" : L" parts found")).c_str ());
					}
				m_count.Invalidate ();
				m_ok.SetWindowText (!m_countOk ? L"Dump"
									: m_more ? L"Dump them all"
									: m_found == 1 ? L"Dump 1 part"
									: (L"Dump " + std::to_wstring (m_found) + L" parts").c_str ());
				m_ok.EnableWindow (m_countOk);
				Where ();
				}

			/// The shop workbook's line: its name, and the folder - the one folder the
			/// workbooks go to, else the folder of parts.
			void Where ()
				{
				if (!m_ready)
					return;				// a box filled in as it is made: not every control is there yet
				const bool one = m_one.GetCheck () == BST_CHECKED;
				m_out.EnableWindow (one);
				m_outBrowse.EnableWindow (one);
				const std::wstring folder = one ? FolderText (m_out) : FolderText (m_folder);
				m_shop.SetWindowText ((L"Then one workbook of every part side by side: " + Shop::FileName (std::time (nullptr))
									   + (folder.empty () ? std::wstring () : L" in " + folder)).c_str ());
				}

			void OnOK () override
				{
				const std::wstring folder = FolderText (m_folder), out = FolderText (m_out);
				const bool one = m_one.GetCheck () == BST_CHECKED;
				std::error_code ec;
				if (!std::filesystem::is_directory (folder, ec))
					{
					MessageBoxW (L"That folder of parts does not exist.", L"Parameter Table Tool", MB_ICONWARNING);
					return;
					}
				if (one && !std::filesystem::is_directory (out, ec))
					{
					MessageBoxW (L"The folder for the workbooks does not exist - pick one, or save them beside each part.",
								 L"Parameter Table Tool", MB_ICONWARNING);
					return;
					}
				m_settings.folder = folder;
				m_settings.subfolders = m_sub.GetCheck () == BST_CHECKED;
				m_settings.outFolder = one ? out : std::wstring ();
				m_settings.pictures = m_pics.GetCheck () == BST_CHECKED;
				m_settings.macros = m_macros.GetCheck () == BST_CHECKED;
				m_settings.stockSim = m_sim.GetCheck () == BST_CHECKED;
				CDialog::OnOK ();
				}

			void Browse (CEdit &e)
				{
				CString folder;
				e.GetWindowText (folder);
				CFolderPickerDialog dlg (folder, 0, this);
				if (dlg.DoModal () == IDOK)
					e.SetWindowText (dlg.GetPathName ());
				}

			afx_msg void OnBrowse ()
				{
				Browse (m_folder);
				Count ();
				}

			afx_msg void OnOutBrowse ()
				{
				Browse (m_out);
				m_one.SetCheck (BST_CHECKED);
				m_beside.SetCheck (BST_UNCHECKED);
				Where ();
				}

			/// Typing in the folder: counted once the typing stops, not at every key.
			afx_msg void OnFolderChange ()
				{
				if (m_ready)
					SetTimer (kCountTimer, 400, nullptr);
				}

			afx_msg void OnTimer (UINT_PTR id)
				{
				if (id == kCountTimer)
					{
					KillTimer (kCountTimer);
					Count ();
					}
				else
					CDialog::OnTimer (id);
				}

			afx_msg void OnSub () { Count (); }
			afx_msg void OnWhere () { Where (); }

			afx_msg HBRUSH OnCtlColor (CDC *dc, CWnd *wnd, UINT ctl)
				{
				HBRUSH br = CDialog::OnCtlColor (dc, wnd, ctl);
				const int id = wnd != nullptr ? wnd->GetDlgCtrlID () : 0;
				if (id == IdWarn)
					dc->SetTextColor (kWarn);
				else if (id == IdShop)
					dc->SetTextColor (kMuted);
				else if (id == IdCount)
					dc->SetTextColor (m_countOk && !m_more ? kGood : kBad);
				else if (id == IdTitle || id == IdPartsHead || id == IdOutHead || id == IdOptHead)
					dc->SetTextColor (kInk);
				return br;
				}

			DECLARE_MESSAGE_MAP ()

		private:
			std::wstring m_openPart, m_ext;
			Settings::Batch &m_settings;
			CWnd *m_parent = nullptr;
			std::vector<WORD> m_tpl;
			Fonts m_f;
			size_t m_found = 0;
			bool m_more = false, m_countOk = false, m_ready = false;

			CStatic m_title, m_warn, m_partsHead, m_count, m_outHead, m_optHead, m_shop;
			CEdit m_folder, m_out;
			CButton m_browse, m_sub, m_beside, m_one, m_outBrowse, m_pics, m_macros, m_sim, m_ok, m_cancel;
		};

	BEGIN_MESSAGE_MAP (Dlg, CDialog)
		ON_BN_CLICKED (IdBrowse, &Dlg::OnBrowse)
		ON_BN_CLICKED (IdOutBrowse, &Dlg::OnOutBrowse)
		ON_EN_CHANGE (IdFolder, &Dlg::OnFolderChange)
		ON_EN_CHANGE (IdOut, &Dlg::OnWhere)
		ON_BN_CLICKED (IdSub, &Dlg::OnSub)
		ON_BN_CLICKED (IdBeside, &Dlg::OnWhere)
		ON_BN_CLICKED (IdOne, &Dlg::OnWhere)
		ON_WM_TIMER ()
		ON_WM_CTLCOLOR ()
	END_MESSAGE_MAP ()

	// ---- The progress window ----------------------------------------------------

	enum { IdStep = 2201, IdName, IdBar, IdTally };

	class ProgressDlg : public CDialog
		{
		public:
			bool cancelled = false;

			bool Open (size_t parts, CWnd *parent)
				{
				m_parts = parts;
				m_tpl.assign (64, 0);
				DLGTEMPLATE *t = reinterpret_cast<DLGTEMPLATE *> (m_tpl.data ());
				t->style = WS_POPUP | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME | WS_VISIBLE;
				t->cx = 200;
				t->cy = 80;
				return CreateIndirect (t, parent) != FALSE;
				}

			void Show (const std::wstring &step, const std::wstring &name, const std::wstring &tally, int at)
				{
				m_step.SetWindowText (step.c_str ());
				m_name.SetWindowText (name.c_str ());
				m_tally.SetWindowText (tally.c_str ());
				if (at >= 0)
					m_bar.SetPos (at);
				UpdateWindow ();
				}

		protected:
			BOOL OnInitDialog () override
				{
				CDialog::OnInitDialog ();
				SetWindowText (L"Parameter Table Tool - dumping a folder of parts");
				m_f.Make (*this);
				auto label = [&] (CStatic &c, int id, CFont &f)
					{
					c.Create (L"", WS_CHILD | WS_VISIBLE | SS_LEFTNOWORDWRAP | SS_NOPREFIX | SS_ENDELLIPSIS, CRect (), this, id);
					c.SetFont (&f);
					};
				label (m_step, IdStep, m_f.bold);
				label (m_name, IdName, m_f.font);
				m_bar.Create (WS_CHILD | WS_VISIBLE | PBS_SMOOTH, CRect (), this, IdBar);
				m_bar.SetRange32 (0, static_cast<int> ((std::max) (static_cast<size_t> (1), m_parts)));
				label (m_tally, IdTally, m_f.font);
				m_cancel.Create (L"Stop after this part", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, CRect (), this,
								 IDCANCEL);
				m_cancel.SetFont (&m_f.font);

				const int u = m_f.u, pad = u, gap = u / 2, w = u * 32, inner = w - 2 * pad;
				const int lineH = u * 3 / 2, btnH = u * 2;
				int y = pad;
				m_step.MoveWindow (pad, y, inner, lineH);
				y += lineH;
				m_name.MoveWindow (pad, y, inner, lineH);
				y += lineH + gap;
				m_bar.MoveWindow (pad, y, inner, u);
				y += u + gap;
				m_tally.MoveWindow (pad, y, inner, lineH);
				y += lineH + gap;
				const int bw = u * 11;
				m_cancel.MoveWindow (w - pad - bw, y, bw, btnH);
				y += btnH + pad;

				CRect work;
				SystemParametersInfo (SPI_GETWORKAREA, 0, &work, 0);
				CRect wr (0, 0, w, y);
				CalcWindowRect (&wr);
				SetWindowPos (&wndTop, work.left + (work.Width () - wr.Width ()) / 2,
							  work.top + (work.Height () - wr.Height ()) / 3, wr.Width (), wr.Height (), 0);
				return TRUE;
				}

			/// Cancel, Escape and the close box all mean: stop after the part in hand.
			/// The window stays until the batch is done with it.
			void OnCancel () override
				{
				if (cancelled)
					return;
				cancelled = true;
				m_cancel.EnableWindow (FALSE);
				m_cancel.SetWindowText (L"Stopping ...");
				m_tally.SetWindowText (L"Stopping after this part - the parts done so far are kept.");
				}

			void OnOK () override
				{
				}

		private:
			size_t m_parts = 0;
			std::vector<WORD> m_tpl;
			Fonts m_f;
			CStatic m_step, m_name, m_tally;
			CProgressCtrl m_bar;
			CButton m_cancel;
		};

	/// Between two parts: the window's own messages (its button, Escape, its close
	/// box), then paints - so Mastercam's window is drawn again after a part opens.
	/// Nothing else of Mastercam's: a timer (an autosave) or a queued click must not
	/// start anything in the middle of the batch.
	void Pump (CWnd &dlg)
		{
		MSG msg;
		const HWND h = dlg.GetSafeHwnd ();
		// A window filter takes its child windows' messages too: the button's click.
		while (h != nullptr && ::PeekMessage (&msg, h, 0, 0, PM_REMOVE))
			if (!::IsDialogMessage (h, &msg))
				{
				::TranslateMessage (&msg);
				::DispatchMessage (&msg);
				}
		// A bounded number: a window that never marks itself painted would come back
		// for ever.
		for (int n = 0; n < 64 && ::PeekMessage (&msg, nullptr, WM_PAINT, WM_PAINT, PM_REMOVE); ++n)
			::DispatchMessage (&msg);
		}
	}

namespace BatchDialog
	{
	bool Show (const std::wstring &openPart, const std::wstring &ext, Settings::Batch &settings)
		{
		Dlg dlg (openPart, ext, settings, Ui::Host ());
		return dlg.Run () == IDOK;
		}

	struct Progress::Impl
		{
		ProgressDlg dlg;
		CWnd *host = nullptr;
		bool hostWasOn = false;
		size_t parts = 0;
		};

	Progress::Progress (size_t parts)
		: m (std::make_unique<Impl> ())
		{
		m->parts = parts;
		m->host = Ui::Host ();
		if (m->host != nullptr && m->host->GetSafeHwnd () != nullptr)
			{
			m->hostWasOn = m->host->IsWindowEnabled () != FALSE;
			m->host->EnableWindow (FALSE);
			}
		m->dlg.Open (parts, m->host);
		Pump (m->dlg);
		}

	Progress::~Progress ()
		{
		// Mastercam's window back on BEFORE this one goes, so it is the one that
		// comes to the front - not whatever program was behind.
		if (m->host != nullptr && m->host->GetSafeHwnd () != nullptr && m->hostWasOn)
			m->host->EnableWindow (TRUE);
		if (m->dlg.GetSafeHwnd () != nullptr)
			m->dlg.DestroyWindow ();
		if (m->host != nullptr && m->host->GetSafeHwnd () != nullptr)
			m->host->SetForegroundWindow ();
		}

	bool Progress::Next (size_t index, const std::wstring &name, size_t done, size_t failed)
		{
		Pump (m->dlg);
		if (m->dlg.cancelled)
			return false;
		if (m->dlg.GetSafeHwnd () != nullptr)
			m->dlg.Show (L"Part " + std::to_wstring (index + 1) + L" of " + std::to_wstring (m->parts), name,
						 std::to_wstring (done) + L" dumped" + (failed > 0 ? L", " + std::to_wstring (failed) + L" could not be" : L""),
						 static_cast<int> (index));
		Pump (m->dlg);
		return !m->dlg.cancelled;
		}

	void Progress::Finishing (const std::wstring &what)
		{
		if (m->dlg.GetSafeHwnd () == nullptr)
			return;
		m->dlg.Show (L"Finishing", what, L"", static_cast<int> (m->parts));
		Pump (m->dlg);
		}
	}
