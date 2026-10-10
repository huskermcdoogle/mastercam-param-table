//
// dialog_shots.cpp - the user manual's pictures of the add-in's own windows,
// made without Mastercam and without a screenshot.
//
// The windows are the add-in's own code - src\DumpDialog.cpp, src\Preview.cpp and
// src\BatchDialog.cpp, which need no Mastercam SDK (src\Ui.h) - shown here with a real part's ops,
// read from a dump workbook that tools\make_help_images.ps1 has edited in Excel
// and saved, and the story the manual's pages tell:
//
//   dump   the dump window: ops found by a word, ticked, the ticks saved
//   load   the load preview of that workbook - the part's time and flips, each
//          op's, every change with its tick box
//   undo   the undo of that same load, one value changed by hand since
//   batch  the batch dump's window, over the part's folder
//
// Each window is moved off the screen as it opens, filled in the way a person
// would (text typed into its boxes, its buttons clicked - messages to its own
// controls), and drawn into a bitmap by ITSELF: PrintWindow on its own handle.
// Nothing on the screen is read. The bitmap goes to <name>.bmp, and where the
// page's numbered callouts go to <name>.txt ("dpi 144", then "<n> <x> <y>" in
// the bitmap's pixels) - the script draws them, as on its other pictures.
//
// The lines of the two previews are put together here the way src\Load.cpp
// does (Run and UndoLast): a change there wants the same change here.
//
// Usage (tools\make_help_images.ps1 builds it with tests\build_dialog_shots.bat
// and runs it):
//
//   dialog_shots <workbook> <output folder> [options]
//     --part <path>               the part file the windows name (its folder is
//                                 the dump's Save to)
//     --edit <op>:<column>        a value the load writes, as dumped -> as on the
//                                 sheet (repeat; in the order the load logs them)
//     --stamp <yyyy-mm-dd hh:mm:ss>  when that load ran
//     --since <op>:<column>=<v>   a value changed by hand after the load
//     --find <words>              typed into the dump window's Find
//     --saved <name>              the matches ticked and saved under this name
//     --only <dump|load|undo|batch>  just one picture (repeat)
//
#include "../src/Ui.h"
#include "../src/Csv.h"
#include "../src/BatchDialog.h"
#include "../src/DumpDialog.h"
#include "../src/Impact.h"
#include "../src/Plan.h"
#include "../src/Preview.h"
#include "../src/Settings.h"
#include "../src/Undo.h"
#include "../src/Xlsx.h"

#include <dwmapi.h>

#include <algorithm>
#include <cstdio>
#include <filesystem>
#include <fstream>
#include <functional>
#include <map>
#include <set>

#pragma comment (lib, "dwmapi.lib")

// The windows as Mastercam shows them: today's themed controls (Common Controls 6).
#pragma comment (linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' "	\
							"version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#ifndef PW_RENDERFULLCONTENT
#define PW_RENDERFULLCONTENT 0x00000002		// Windows 8.1 and later; the SDK is pinned to Windows 7
#endif

CWinApp theApp;			// MFC's windows need an application object, even here

// ---- What the windows need from outside them, without Mastercam --------------

namespace Ui
	{
	/// No Mastercam window to sit over: the windows have no owner.
	CWnd *Host ()
		{
		return nullptr;
		}
	}

namespace
	{
	/// The saved ticks, here in memory - the add-in keeps them in the registry,
	/// and a picture must not touch the settings of whoever makes it.
	std::vector<std::pair<std::wstring, std::vector<long>>> gSaved;
	}

namespace Settings
	{
	std::vector<std::pair<std::wstring, std::vector<long>>> Selections (const std::wstring &)
		{
		return gSaved;
		}

	void SaveSelection (const std::wstring &, const std::wstring &name, const std::vector<long> &ids)
		{
		for (auto &s : gSaved)
			if (_wcsicmp (s.first.c_str (), name.c_str ()) == 0)
				{
				s.second = ids;
				return;
				}
		gSaved.push_back ({ name, ids });
		std::sort (gSaved.begin (), gSaved.end ());
		}

	void DeleteSelection (const std::wstring &, const std::wstring &name)
		{
		gSaved.erase (std::remove_if (gSaved.begin (), gSaved.end (),
									  [&] (const auto &s) { return _wcsicmp (s.first.c_str (), name.c_str ()) == 0; }),
					  gSaved.end ());
		}
	}

namespace
	{
	// ---- The sheet -------------------------------------------------------------

	/// A sheet as read back: the column names, the rows, each row's place in Excel.
	struct Sheet
		{
		std::vector<Csv::Row> rows;			// [0] = the column names
		std::vector<size_t> rowNo;

		int Col (const std::wstring &name) const
			{
			if (rows.empty ())
				return -1;
			for (size_t i = 0; i < rows[0].size (); ++i)
				if (Csv::Trim (rows[0][i]) == name)
					return static_cast<int> (i);
			return -1;
			}

		/// The row of an op, or 0 (the column-name row) when it has none.
		size_t RowOf (long op) const
			{
			for (size_t k = 1; k < rows.size (); ++k)
				{
				long long id = 0;
				if (!rows[k].empty () && Csv::ParseLong (rows[k][0], id) && id == op)
					return k;
				}
			return 0;
			}

		std::wstring Cell (size_t row, int col) const
			{
			if (row >= rows.size () || col < 0 || static_cast<size_t> (col) >= rows[row].size ())
				return std::wstring ();
			return rows[row][static_cast<size_t> (col)];
			}

		std::wstring Cell (long op, const std::wstring &col) const
			{
			const size_t r = RowOf (op);
			return r == 0 ? std::wstring () : Cell (r, Col (col));
			}
		};

	/// One value the load writes.
	struct Edit
		{
		long op = 0;
		std::wstring column;
		std::wstring type;
		std::wstring from;			// as dumped
		std::wstring to;			// as on the sheet
		size_t row = 0;				// its row in Excel
		int col = 0;				// its column on the sheet
		};

	std::wstring Name (const std::filesystem::path &p)
		{
		return p.filename ().wstring ();
		}

	Preview::Line MakeLine (Preview::Line::Kind k, const std::wstring &text,
							const std::wstring &detail = std::wstring (),
							const std::wstring &from = std::wstring (),
							const std::wstring &to = std::wstring ())
		{
		Preview::Line l;
		l.kind = k;
		l.text = text;
		l.detail = detail;
		l.from = from;
		l.to = to;
		return l;
		}

	// ---- Showing a window and drawing it ------------------------------------------

	/// One numbered callout, in the picture's pixels.
	struct Callout
		{
		int n = 0;
		int x = 0, y = 0;
		};

	/// The picture being made: what to do to the window once it is up, and where
	/// its callouts go.
	struct Scene
		{
		std::wstring name;
		std::function<void (HWND)> fill;
		std::function<std::vector<Callout> (HWND, const RECT &frame, int dpi)> callouts;
		};

	Scene *gScene = nullptr;
	HWND gDlg = nullptr;
	int gStep = 0, gWaited = 0;
	UINT_PTR gTimer = 0;
	std::filesystem::path gOut;
	bool gMade = false;
	HHOOK gHook = nullptr;

	/// Somewhere no screen is: left of every monitor.
	POINT OffScreen (HWND h)
		{
		RECT r;
		GetWindowRect (h, &r);
		return { GetSystemMetrics (SM_XVIRTUALSCREEN) - (r.right - r.left) - 2000,
				 GetSystemMetrics (SM_YVIRTUALSCREEN) };
		}

	/// As the window finishes setting itself up (sized, centred), it is moved
	/// off the screen - before it is shown, so it never appears on one.
	LRESULT CALLBACK AfterMessage (int code, WPARAM w, LPARAM l)
		{
		const CWPRETSTRUCT *m = reinterpret_cast<const CWPRETSTRUCT *> (l);
		if (code == HC_ACTION && m != nullptr && m->message == WM_INITDIALOG && gScene != nullptr
			&& gDlg == nullptr && GetParent (m->hwnd) == nullptr)
			{
			gDlg = m->hwnd;
			const POINT at = OffScreen (gDlg);
			SetWindowPos (gDlg, nullptr, at.x, at.y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
			}
		return CallNextHookEx (gHook, code, w, l);
		}

	int Dpi ()
		{
		HDC dc = GetDC (nullptr);
		const int dpi = GetDeviceCaps (dc, LOGPIXELSY);
		ReleaseDC (nullptr, dc);
		return dpi;
		}

	/// The window's frame as you see it (no invisible resize border), on the screen.
	RECT Frame (HWND h)
		{
		RECT f;
		if (FAILED (DwmGetWindowAttribute (h, DWMWA_EXTENDED_FRAME_BOUNDS, &f, sizeof (f))))
			GetWindowRect (h, &f);
		return f;
		}

	/// A 24-bit BMP: rows bottom up, each padded to four bytes.
	bool SaveBmp (const std::filesystem::path &file, const BYTE *bgra, int stride, int x0, int y0, int w, int h)
		{
		const int row = (w * 3 + 3) & ~3;
		BITMAPFILEHEADER fh = {};
		BITMAPINFOHEADER ih = {};
		fh.bfType = 0x4D42;
		fh.bfOffBits = sizeof (fh) + sizeof (ih);
		fh.bfSize = fh.bfOffBits + static_cast<DWORD> (row) * h;
		ih.biSize = sizeof (ih);
		ih.biWidth = w;
		ih.biHeight = h;
		ih.biPlanes = 1;
		ih.biBitCount = 24;
		ih.biCompression = BI_RGB;
		std::ofstream out (file, std::ios::binary);
		if (!out)
			return false;
		out.write (reinterpret_cast<const char *> (&fh), sizeof (fh));
		out.write (reinterpret_cast<const char *> (&ih), sizeof (ih));
		std::vector<char> line (static_cast<size_t> (row), 0);
		for (int y = h - 1; y >= 0; --y)
			{
			const BYTE *src = bgra + static_cast<size_t> (y0 + y) * stride + static_cast<size_t> (x0) * 4;
			for (int x = 0; x < w; ++x)
				{
				line[static_cast<size_t> (x) * 3 + 0] = static_cast<char> (src[x * 4 + 0]);
				line[static_cast<size_t> (x) * 3 + 1] = static_cast<char> (src[x * 4 + 1]);
				line[static_cast<size_t> (x) * 3 + 2] = static_cast<char> (src[x * 4 + 2]);
				}
			out.write (line.data (), row);
			}
		return static_cast<bool> (out);
		}

	/// The window drawn by itself (PrintWindow, its full content) and cut to its
	/// visible frame; the callouts beside it.
	bool Shoot (HWND h)
		{
		RECT wr;
		GetWindowRect (h, &wr);
		const RECT f = Frame (h);
		const int ww = wr.right - wr.left, wh = wr.bottom - wr.top;
		BITMAPINFO bi = {};
		bi.bmiHeader.biSize = sizeof (bi.bmiHeader);
		bi.bmiHeader.biWidth = ww;
		bi.bmiHeader.biHeight = -wh;			// top down
		bi.bmiHeader.biPlanes = 1;
		bi.bmiHeader.biBitCount = 32;
		bi.bmiHeader.biCompression = BI_RGB;
		void *bits = nullptr;
		HDC dc = CreateCompatibleDC (nullptr);
		HBITMAP bmp = CreateDIBSection (dc, &bi, DIB_RGB_COLORS, &bits, nullptr, 0);
		if (bmp == nullptr || bits == nullptr)
			{
			DeleteDC (dc);
			return false;
			}
		HGDIOBJ old = SelectObject (dc, bmp);
		const BOOL printed = PrintWindow (h, dc, PW_RENDERFULLCONTENT);
		GdiFlush ();
		bool ok = false;
		if (printed)
			{
			RECT at = f;				// the picture's top left, on the screen
			int x0 = f.left - wr.left, y0 = f.top - wr.top, w = f.right - f.left, ht = f.bottom - f.top;
			if (x0 < 0 || y0 < 0 || w <= 0 || ht <= 0 || x0 + w > ww || y0 + ht > wh)
				{
				at = wr;
				x0 = y0 = 0;
				w = ww;
				ht = wh;
				}
			ok = SaveBmp (gOut / (gScene->name + L".bmp"), static_cast<const BYTE *> (bits), ww * 4, x0, y0, w, ht);
			const int dpi = Dpi ();
			std::ofstream places (gOut / (gScene->name + L".txt"));
			places << "dpi " << dpi << "\n";
			if (gScene->callouts)
				for (const Callout &c : gScene->callouts (h, at, dpi))
					places << c.n << " " << c.x << " " << c.y << "\n";
			}
		SelectObject (dc, old);
		DeleteObject (bmp);
		DeleteDC (dc);
		return ok;
		}

	/// The window's own clock, in its own modal loop: fill it in, give it time to
	/// draw, draw it, close it (Cancel - nothing is ever applied). A window that
	/// never shows is closed after a while, so the script running this never waits
	/// for ever.
	void CALLBACK Tick (HWND, UINT, UINT_PTR, DWORD)
		{
		if (gDlg == nullptr || !IsWindowVisible (gDlg))
			{
			if (++gWaited > 80 && gDlg != nullptr)
				{
				KillTimer (nullptr, gTimer);
				gTimer = 0;
				PostMessage (gDlg, WM_COMMAND, MAKEWPARAM (IDCANCEL, BN_CLICKED), 0);
				}
			return;
			}
		++gStep;
		if (gStep == 1 && gScene->fill)
			gScene->fill (gDlg);
		if (gStep < 4)
			return;
		KillTimer (nullptr, gTimer);
		gTimer = 0;
		RedrawWindow (gDlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW);
		gMade = Shoot (gDlg);
		PostMessage (gDlg, WM_COMMAND, MAKEWPARAM (IDCANCEL, BN_CLICKED), 0);
		}

	/// Show a window (`show` opens it and returns when it closes) as `scene`.
	bool Make (Scene &scene, const std::function<void ()> &show)
		{
		gScene = &scene;
		gDlg = nullptr;
		gStep = gWaited = 0;
		gMade = false;
		gHook = SetWindowsHookEx (WH_CALLWNDPROCRET, AfterMessage, nullptr, GetCurrentThreadId ());
		gTimer = SetTimer (nullptr, 0, 250, Tick);
		show ();
		if (gTimer != 0)
			KillTimer (nullptr, gTimer);
		gTimer = 0;
		UnhookWindowsHookEx (gHook);
		gHook = nullptr;
		gScene = nullptr;
		std::printf ("  %s  %ls\n", gMade ? "ok   " : "FAIL ", scene.name.c_str ());
		return gMade;
		}

	// ---- Finding a window's controls, from outside it ------------------------------

	std::wstring TextOf (HWND h)
		{
		wchar_t buf[1024] = L"";
		GetWindowText (h, buf, 1024);
		return buf;
		}

	std::wstring ClassOf (HWND h)
		{
		wchar_t buf[128] = L"";
		GetClassName (h, buf, 128);
		return buf;
		}

	/// The window's own controls (not a combo box's inner edit) that `want` picks.
	std::vector<HWND> Controls (HWND dlg, const std::function<bool (HWND)> &want)
		{
		std::vector<HWND> all;
		for (HWND c = GetWindow (dlg, GW_CHILD); c != nullptr; c = GetWindow (c, GW_HWNDNEXT))
			if (want (c))
				all.push_back (c);
		return all;
		}

	HWND Button (HWND dlg, const std::wstring &text, bool prefix = false)
		{
		const auto found = Controls (dlg, [&] (HWND c)
			{
			const std::wstring t = TextOf (c);
			return ClassOf (c) == L"Button" && (prefix ? t.compare (0, text.size (), text) == 0 : t == text);
			});
		return found.empty () ? nullptr : found.front ();
		}

	void Click (HWND dlg, HWND button)
		{
		if (button != nullptr)
			SendMessage (dlg, WM_COMMAND, MAKEWPARAM (GetDlgCtrlID (button), BN_CLICKED),
						 reinterpret_cast<LPARAM> (button));
		}

	/// A control's place in the picture (whose top left is the frame's).
	RECT InPicture (HWND c, const RECT &frame)
		{
		RECT r;
		GetWindowRect (c, &r);
		OffsetRect (&r, -frame.left, -frame.top);
		return r;
		}

	/// How wide `text` is in the font a control draws with (or `font`).
	int TextW (HWND c, const std::wstring &text, HFONT font = nullptr)
		{
		HDC dc = GetDC (c);
		if (font == nullptr)
			font = reinterpret_cast<HFONT> (SendMessage (c, WM_GETFONT, 0, 0));
		HGDIOBJ old = SelectObject (dc, font);
		SIZE s = {};
		GetTextExtentPoint32 (dc, text.c_str (), static_cast<int> (text.size ()), &s);
		SelectObject (dc, old);
		ReleaseDC (c, dc);
		return s.cx;
		}

	int TextH (HWND c)
		{
		HDC dc = GetDC (c);
		HGDIOBJ old = SelectObject (dc, reinterpret_cast<HFONT> (SendMessage (c, WM_GETFONT, 0, 0)));
		TEXTMETRIC tm = {};
		GetTextMetrics (dc, &tm);
		SelectObject (dc, old);
		ReleaseDC (c, dc);
		return tm.tmHeight;
		}

	/// The callout's size, as the script draws it: 17 points across.
	int CalloutD (int dpi)
		{
		return 17 * dpi / 72;
		}

	/// A callout just right of `x` on a line through `y`.
	Callout RightOf (int n, int x, int y, int dpi)
		{
		const int d = CalloutD (dpi);
		return { n, x + d / 4 + d / 2, y };
		}

	/// A callout just left of a control, on its middle.
	Callout LeftOf (int n, const RECT &r, int dpi)
		{
		const int d = CalloutD (dpi);
		return { n, (std::max) (d / 2, static_cast<int> (r.left) - d / 4 - d / 2), (r.top + r.bottom) / 2 };
		}

	/// A callout inside a box, at its right end.
	Callout InsideRight (int n, const RECT &r, int dpi)
		{
		const int d = CalloutD (dpi);
		return { n, r.right - d / 4 - d / 2, (r.top + r.bottom) / 2 };
		}

	// ---- The list of a preview -------------------------------------------------------

	HWND ListOf (HWND dlg)
		{
		const auto l = Controls (dlg, [] (HWND c) { return ClassOf (c) == WC_LISTVIEW; });
		return l.empty () ? nullptr : l.front ();
		}

	/// Line `item` of the list (a part of it: LVIR_BOUNDS the whole line, or a
	/// column's cell), in the picture.
	RECT LineInPicture (HWND list, int item, int column, const RECT &frame)
		{
		RECT r = {};
		if (column == 0)
			ListView_GetItemRect (list, item, &r, LVIR_BOUNDS);
		else
			ListView_GetSubItemRect (list, item, column, LVIR_LABEL, &r);
		MapWindowPoints (list, nullptr, reinterpret_cast<POINT *> (&r), 2);
		OffsetRect (&r, -frame.left, -frame.top);
		return r;
		}

	/// The window as tall as its lines need (and a line to spare), the way a
	/// person would size it - a short list leaves no empty half window.
	void FitToLines (HWND dlg)
		{
		HWND list = ListOf (dlg);
		const int n = ListView_GetItemCount (list);
		if (list == nullptr || n == 0)
			return;
		RECT last = {}, client = {};
		ListView_GetItemRect (list, n - 1, &last, LVIR_BOUNDS);
		GetClientRect (list, &client);
		const int spare = client.bottom - (last.bottom + (last.bottom - last.top));
		if (spare <= 0)
			return;
		RECT w;
		GetWindowRect (dlg, &w);
		SetWindowPos (dlg, nullptr, 0, 0, w.right - w.left, w.bottom - w.top - spare,
					  SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
		}

	/// Where line `item`'s tick box is drawn: as src\Preview.cpp's BoxRect.
	RECT BoxInPicture (HWND list, int item, bool change, const RECT &frame)
		{
		const int u = (std::max) (12, TextH (list));
		const int box = (std::max) (11, u * 4 / 5);
		RECT r = LineInPicture (list, item, 0, frame);
		const int left = r.left + 8 + (change ? u : 0);
		const int top = r.top + (r.bottom - r.top - box) / 2;
		return { left, top, left + box, top + box };
		}

	// ---- Options -----------------------------------------------------------------

	struct Args
		{
		std::filesystem::path workbook, out;
		std::wstring part = L"C:\\Parts\\Oil Spool\\Oil Spool Head_2026.mcam";
		std::vector<std::pair<long, std::wstring>> edits;
		std::wstring stamp;
		std::map<std::pair<long, std::wstring>, std::wstring> since;
		std::wstring find, saved;
		std::set<std::wstring> only;
		std::map<long, std::wstring> comments;	// each op's, from the sheet - as the windows name ops
		};

	/// "<op>:<column>" (and "=<value>" after it, when `value` is given).
	bool OpColumn (const std::wstring &s, long &op, std::wstring &column, std::wstring *value)
		{
		const size_t colon = s.find (L':');
		if (colon == std::wstring::npos)
			return false;
		long long id = 0;
		if (!Csv::ParseLong (s.substr (0, colon), id))
			return false;
		op = static_cast<long> (id);
		column = s.substr (colon + 1);
		if (value != nullptr)
			{
			const size_t eq = column.find (L'=');
			if (eq == std::wstring::npos)
				return false;
			*value = column.substr (eq + 1);
			column = column.substr (0, eq);
			}
		return !column.empty ();
		}

	bool ReadArgs (int argc, wchar_t **argv, Args &a)
		{
		if (argc < 3)
			return false;
		a.workbook = argv[1];
		a.out = argv[2];
		for (int i = 3; i + 1 < argc; i += 2)
			{
			const std::wstring k = argv[i], v = argv[i + 1];
			long op = 0;
			std::wstring col, val;
			if (k == L"--part")
				a.part = v;
			else if (k == L"--edit" && OpColumn (v, op, col, nullptr))
				a.edits.push_back ({ op, col });
			else if (k == L"--stamp")
				a.stamp = v;
			else if (k == L"--since" && OpColumn (v, op, col, &val))
				a.since[{ op, col }] = val;
			else if (k == L"--find")
				a.find = v;
			else if (k == L"--saved")
				a.saved = v;
			else if (k == L"--only")
				a.only.insert (v);
			else
				{
				std::printf ("dialog_shots: what is \"%ls %ls\"?\n", k.c_str (), v.c_str ());
				return false;
				}
			}
		return true;
		}

	// ---- The pictures ----------------------------------------------------------------

	/// The dump window: every op of the sheet in its toolpath group, a search typed,
	/// None, Tick matches, the ticks saved under a name.
	bool DumpWindow (const Args &a, const Sheet &sheet)
		{
		const int cType = sheet.Col (L"type"), cTool = sheet.Col (L"tool"), cComment = sheet.Col (L"comment"),
				  cGroup = sheet.Col (L"group_name");
		std::vector<DumpDialog::Op> ops;
		std::map<std::wstring, long> groups;
		for (size_t k = 1; k < sheet.rows.size (); ++k)
			{
			long long id = 0;
			if (!Csv::ParseLong (sheet.Cell (k, 0), id))
				continue;
			// As src\Dump.cpp words an op: number, kind, tool (the kinds that have one), comment.
			DumpDialog::Op o;
			o.idn = static_cast<long> (id);
			o.groupName = sheet.Cell (k, cGroup);
			if (groups.count (o.groupName) == 0)
				groups[o.groupName] = static_cast<long> (groups.size ()) + 1;
			o.group = groups[o.groupName];
			o.kind = sheet.Cell (k, cType);
			const std::wstring tool = sheet.Cell (k, cTool);
			o.text = L"op " + std::to_wstring (o.idn) + L"    " + o.kind
					 + (!tool.empty () && tool != L"0" ? L"    T" + tool : std::wstring ())
					 + L"    " + sheet.Cell (k, cComment);
			o.on = true;
			ops.push_back (o);
			}

		// Its Save to is the part's folder ("Beside the part"); the window checks the
		// folder is there, so one that is not is made for the picture and taken away after.
		const std::filesystem::path folder = std::filesystem::path (a.part).parent_path ();
		std::vector<std::filesystem::path> made;
		std::error_code ec;
		for (std::filesystem::path p = folder; !p.empty () && !std::filesystem::exists (p, ec); p = p.parent_path ())
			made.push_back (p);
		if (!made.empty ())
			std::filesystem::create_directories (folder, ec);

		Settings::Dump settings;
		settings.macros = true;

		Scene scene;
		scene.name = L"dump";
		scene.fill = [&a] (HWND dlg)
			{
			const auto edits = Controls (dlg, [] (HWND c) { return ClassOf (c) == L"Edit"; });
			// The find box is the one edit box that starts empty (Save to and the file name do not).
			for (HWND e : edits)
				if (TextOf (e).empty ())
					SetWindowText (e, a.find.c_str ());
			Click (dlg, Button (dlg, L"None"));
			Click (dlg, Button (dlg, L"Tick matches"));
			if (!a.saved.empty ())
				{
				const auto combos = Controls (dlg, [] (HWND c)
					{
					return ClassOf (c) == L"ComboBox" && (GetWindowLong (c, GWL_STYLE) & 3) == CBS_DROPDOWN;
					});
				if (!combos.empty ())
					SetWindowText (combos.front (), a.saved.c_str ());
				Click (dlg, Button (dlg, L"Save ticks"));
				}
			// The list scrolled to its top, so its group line shows: the window opens
			// with it one line down (as it does in Mastercam).
			for (HWND t : Controls (dlg, [] (HWND c) { return ClassOf (c) == WC_TREEVIEW; }))
				SendMessage (t, WM_VSCROLL, MAKEWPARAM (SB_TOP, 0), 0);
			};
		scene.callouts = [&a] (HWND dlg, const RECT &frame, int dpi)
			{
			std::vector<Callout> c;
			const auto trees = Controls (dlg, [] (HWND x) { return ClassOf (x) == WC_TREEVIEW; });
			if (!trees.empty ())
				{
				// On its top line (the group), at the right of the list.
				RECT t = InPicture (trees.front (), frame);
				const int row = TreeView_GetItemHeight (trees.front ());
				const int bar = (GetWindowLong (trees.front (), GWL_STYLE) & WS_VSCROLL) ? GetSystemMetrics (SM_CXVSCROLL) : 0;
				t.right -= bar;
				t.top += 2;
				t.bottom = t.top + row;
				c.push_back (InsideRight (1, t, dpi));
				}
			for (HWND e : Controls (dlg, [] (HWND x) { return ClassOf (x) == L"Edit"; }))
				{
				const std::wstring text = TextOf (e);
				const int n = text == a.find ? 2 : text.find (L'{') != std::wstring::npos ? 4 : 3;
				c.push_back (InsideRight (n, InPicture (e, frame), dpi));
				}
			if (HWND m = Button (dlg, L"Include macros", true))
				{
				// After its words: the box and its gap, then the text.
				const RECT r = InPicture (m, frame);
				const int box = GetSystemMetrics (SM_CXMENUCHECK) + TextH (m) / 2;
				c.push_back (RightOf (5, r.left + box + TextW (m, TextOf (m)), (r.top + r.bottom) / 2, dpi));
				}
			c.push_back (LeftOf (6, InPicture (GetDlgItem (dlg, IDOK), frame), dpi));
			std::sort (c.begin (), c.end (), [] (const Callout &x, const Callout &y) { return x.n < y.n; });
			return c;
			};

		Settings::Dump chosen = settings;
		std::vector<DumpDialog::Op> shown = ops;
		const bool ok = Make (scene, [&] { DumpDialog::Show (a.part, shown, std::wstring (), chosen); });

		for (const std::filesystem::path &p : made)
			std::filesystem::remove (p, ec);			// only if still empty
		return ok;
		}

	/// The batch window: the part's folder as the folder of parts, with a few parts in
	/// it (any that are not there are made, empty, for the picture and taken away
	/// after - so the count reads as on a real folder), each workbook beside its
	/// part, the options as they start.
	bool BatchWindow (const Args &a)
		{
		const std::filesystem::path folder = std::filesystem::path (a.part).parent_path ();
		std::vector<std::filesystem::path> madeDirs, madeFiles;
		std::error_code ec;
		for (std::filesystem::path p = folder; !p.empty () && !std::filesystem::exists (p, ec); p = p.parent_path ())
			madeDirs.push_back (p);
		if (!madeDirs.empty ())
			std::filesystem::create_directories (folder, ec);
		for (const std::wstring &name : { std::filesystem::path (a.part).filename ().wstring (), std::wstring (L"Bushing.mcam"),
										  std::wstring (L"Cap.mcam"), std::wstring (L"Flange.mcam"), std::wstring (L"Shaft.mcam"),
										  std::wstring (L"Sleeve.mcam") })
			{
			const std::filesystem::path f = folder / name;
			if (!std::filesystem::exists (f, ec))
				{
				std::ofstream (f) << "";
				madeFiles.push_back (f);
				}
			}

		Settings::Batch settings;
		settings.folder = folder.wstring ();

		Scene scene;
		scene.name = L"batch";
		scene.callouts = [] (HWND dlg, const RECT &frame, int dpi)
			{
			std::vector<Callout> c;
			// After a static's (or a tick box's) words: its box first, for a button.
			auto after = [&] (int n, HWND h, bool box)
				{
				if (h == nullptr)
					return;
				const RECT r = InPicture (h, frame);
				const int at = r.left + (box ? GetSystemMetrics (SM_CXMENUCHECK) + TextH (h) / 2 : 0) + TextW (h, TextOf (h));
				c.push_back (RightOf (n, at, box ? (r.top + r.bottom) / 2 : r.top + TextH (h) / 2, dpi));
				};
			for (HWND s : Controls (dlg, [] (HWND x) { return ClassOf (x) == L"Static"; }))
				{
				const std::wstring t = TextOf (s);
				if (t == L"Dump a folder of parts")
					after (1, s, false);
				else if (t.find (L"found") != std::wstring::npos)
					after (3, s, false);
				}
			const auto edits = Controls (dlg, [] (HWND x) { return ClassOf (x) == L"Edit"; });
			if (!edits.empty ())
				c.push_back (InsideRight (2, InPicture (edits.front (), frame), dpi));
			after (4, Button (dlg, L"Beside each part", true), true);
			after (5, Button (dlg, L"Simulate the stock", true), true);
			c.push_back (LeftOf (6, InPicture (GetDlgItem (dlg, IDOK), frame), dpi));
			std::sort (c.begin (), c.end (), [] (const Callout &x, const Callout &y) { return x.n < y.n; });
			return c;
			};
		const bool ok = Make (scene, [&] { BatchDialog::Show (a.part, L".mcam", settings); });

		for (const std::filesystem::path &f : madeFiles)
			std::filesystem::remove (f, ec);
		for (const std::filesystem::path &p : madeDirs)
			std::filesystem::remove (p, ec);			// only if still empty
		return ok;
		}

	/// The load's lines, as src\Load.cpp's Run puts them together.
	bool LoadPreview (const Args &a, const Sheet &sheet, const std::vector<Csv::Row> &dumped,
					  const std::vector<Edit> &edits)
		{
		const std::map<long, Impact::Op> figures = Impact::Read (sheet.rows, dumped);
		std::map<long, std::wstring> comments = a.comments;

		// Changes by operation, in sheet order; an op's changes in its columns' order.
		std::vector<Edit> items = edits;
		std::stable_sort (items.begin (), items.end (), [] (const Edit &x, const Edit &y)
			{ return x.row != y.row ? x.row < y.row : x.col < y.col; });
		std::set<long> changedOps;
		for (const Edit &e : items)
			changedOps.insert (e.op);
		const int changes = static_cast<int> (items.size ()), rowsChanged = static_cast<int> (changedOps.size ());
		int unchanged = 0;
		for (size_t k = 1; k < sheet.rows.size (); ++k)
			{
			long long id = 0;
			if (Csv::ParseLong (sheet.Cell (k, 0), id) && changedOps.count (static_cast<long> (id)) == 0)
				++unchanged;
			}

		std::vector<Preview::Line> lines;
		lines.push_back (MakeLine (Preview::Line::Section, L"Changes - " + std::to_wstring (changes)
								   + (changes == 1 ? L" value on " : L" values on ") + std::to_wstring (rowsChanged)
								   + (rowsChanged == 1 ? L" operation" : L" operations")));
		long lastOp = -1;
		int opLine = -1, changeLine = -1;
		for (size_t k = 0; k < items.size (); ++k)
			{
			const Edit &it = items[k];
			if (it.op != lastOp)
				{
				Preview::Line o = MakeLine (Preview::Line::Op, L"op " + std::to_wstring (it.op) + L"  \u00B7  " + it.type
											+ L"  \u00B7  row " + std::to_wstring (it.row), comments[it.op]);
				o.box = Preview::Line::Ticked;
				o.tag = it.op;
				const auto f = figures.find (it.op);
				if (f != figures.end ())
					{
					o.impact = Impact::OpText (f->second);
					o.tone = Impact::Tone (f->second.was, f->second.now);
					if (o.tone == 0)
						o.tone = Impact::Tone (f->second.flipsWas, f->second.flipsNow, true);
					}
				if (opLine < 0 && !o.impact.empty ())
					opLine = static_cast<int> (lines.size ());
				lines.push_back (o);
				lastOp = it.op;
				}
			Preview::Line c = MakeLine (Preview::Line::Change, it.column, std::wstring (), it.from, it.to);
			c.box = Preview::Line::Ticked;
			c.link = Plan::LinkOf (it.column);
			c.tag = static_cast<long> (k);
			if (changeLine < 0 && opLine >= 0)
				changeLine = static_cast<int> (lines.size ());
			lines.push_back (c);
			}

		const std::wstring file = Name (a.workbook);
		const std::wstring summary =
			std::to_wstring (changes) + (changes == 1 ? L" change" : L" changes")
			+ L"  \u00B7  " + std::to_wstring (rowsChanged) + L" operation(s) changed"
			+ L"  \u00B7  " + std::to_wstring (unchanged) + L" unchanged";

		Preview::Options options;
		options.foot = L"Untick a change to leave it out. Every old value goes to ParamTable.log first - "
					   L"\"Lathe params - undo last load\" puts them back. Changed operations are marked "
					   L"for regeneration.";
		if (!figures.empty ())
			options.impact = [&figures] (std::vector<Preview::Line> &ls, int &tone) -> std::wstring
				{
				std::set<long> applied;
				bool partly = false;
				for (const Preview::Line &l : ls)
					if (l.kind == Preview::Line::Op && l.box != Preview::Line::Unticked)
						{
						applied.insert (l.tag);
						partly = partly || l.box == Preview::Line::Mixed;
						}
				const Impact::Total t = Impact::Sum (figures, applied);
				std::wstring s = Impact::TotalText (t);
				if (s.empty ())
					return s;
				tone = Impact::Tone (t.was, t.now);
				if (tone == 0)
					tone = Impact::Tone (t.flipsWas, t.flipsNow, true);
				if (partly)
					s += L"  (partly ticked operations count all their edits)";
				return s;
				};
		bool regenAfter = false;
		options.choice = [] (const std::vector<Preview::Line> &ls)
			{
			int n = 0;
			for (const Preview::Line &l : ls)
				if (l.kind == Preview::Line::Op && l.box != Preview::Line::Unticked)
					++n;
			return L"Regenerate the " + std::to_wstring (n) + (n == 1 ? L" changed operation" : L" changed operations")
				   + L" after loading - can take a long time on big parts";
			};
		options.chosen = &regenAfter;

		Scene scene;
		scene.name = L"load";
		scene.fill = FitToLines;
		scene.callouts = [&lines, opLine, changeLine] (HWND dlg, const RECT &frame, int dpi)
			{
			std::vector<Callout> c;
			HWND list = ListOf (dlg);
			HWND apply = GetDlgItem (dlg, IDOK);
			const HFONT bold = reinterpret_cast<HFONT> (SendMessage (apply, WM_GETFONT, 0, 0));
			// 1: after the part's line.
			for (HWND s : Controls (dlg, [] (HWND x) { return ClassOf (x) == L"Static"; }))
				{
				const std::wstring text = TextOf (s);
				if (text.compare (0, 10, L"Cycle time") == 0)
					{
					// On its line of text, which sits at the top of a taller label.
					const RECT r = InPicture (s, frame);
					c.push_back (RightOf (1, r.left + TextW (s, text), r.top + TextH (s) / 2, dpi));
					}
				}
			// 2: before an op's time, at the right of its line.
			if (list != nullptr && opLine >= 0)
				{
				const RECT r = LineInPicture (list, opLine, 0, frame);
				const int x = r.right - 8 - TextW (list, lines[static_cast<size_t> (opLine)].impact, bold);
				c.push_back ({ 2, x - CalloutD (dpi) / 4 - CalloutD (dpi) / 2, (r.top + r.bottom) / 2 });
				}
			// 3: after a change's new value; 4: beside its tick box.
			if (list != nullptr && changeLine >= 0)
				{
				const RECT r = LineInPicture (list, changeLine, 3, frame);
				c.push_back (RightOf (3, r.left + 6 + TextW (list, lines[static_cast<size_t> (changeLine)].to, bold),
									  (r.top + r.bottom) / 2, dpi));
				const RECT box = BoxInPicture (list, changeLine, true, frame);
				c.push_back (RightOf (4, box.right, (box.top + box.bottom) / 2, dpi));
				}
			c.push_back (LeftOf (5, InPicture (apply, frame), dpi));
			return c;
			};
		return Make (scene, [&] { Preview::Show (L"Load " + file, summary, lines, options); });
		}

	/// The undo of that load, as src\Load.cpp's UndoLast: the log it wrote, read
	/// back by Undo (the add-in's own code), against the ops as they are now.
	bool UndoWindow (const Args &a, const std::vector<Edit> &edits)
		{
		const std::wstring part = Name (a.part), file = Name (a.workbook);
		std::set<long> ops;
		std::wstring log = a.stamp + L"  " + Undo::BeginLine (part, file) + L"\r\n";
		for (const Edit &e : edits)
			{
			log += a.stamp + L"  " + Undo::ChangeLine (e.op, e.type, e.column, e.from, e.to) + L"\r\n";
			ops.insert (e.op);
			}
		log += a.stamp + L"  load: " + std::to_wstring (ops.size ()) + L" operation(s) written, 0 failed\r\n";
		const Undo::Last last = Undo::FindLast (log, part);

		// Each kind's columns (the ones the load wrote) and its ops as they are now:
		// as the load left them, but for the values changed by hand since.
		std::map<std::wstring, Plan::Schema> schemas;
		std::map<std::wstring, std::map<long, Plan::Current>> now;
		for (const Edit &e : edits)
			{
			Plan::Schema &s = schemas[e.type];
			s.type = e.type;
			if (s.Find (e.column) < 0)
				{
				Plan::Col c;
				c.name = e.column;
				double v = 0;
				c.type = Csv::ParseDouble (e.from, v) || Csv::ParseDouble (e.to, v) ? Plan::Type::Double : Plan::Type::Text;
				s.cols.push_back (c);
				}
			}
		for (const Edit &e : edits)
			{
			const Plan::Schema &s = schemas[e.type];
			Plan::Current &c = now[e.type][e.op];
			c.op = e.op;
			c.type = e.type;
			c.values.resize (s.cols.size ());
			const auto hand = a.since.find ({ e.op, e.column });
			c.values[static_cast<size_t> (s.Find (e.column))] = hand != a.since.end () ? hand->second : e.to;
			}
		std::vector<Undo::Kind> kinds;
		for (const auto &kv : schemas)
			{
			Undo::Kind k;
			k.schema = &kv.second;
			for (const auto &c : now[kv.first])
				k.now.push_back (c.second);
			kinds.push_back (k);
			}
		const Undo::Result r = Undo::Make (last, kinds);

		std::map<long, std::wstring> comments = a.comments;
		std::vector<Preview::Line> lines;
		int sinceLine = -1, changeLine = -1;
		if (!r.skipped.empty ())
			{
			sinceLine = static_cast<int> (lines.size ());
			lines.push_back (MakeLine (Preview::Line::Section, L"Changed since the load - left as they are ("
									   + std::to_wstring (r.skipped.size ()) + L")"));
			for (const Undo::Skip &s : r.skipped)
				lines.push_back (MakeLine (Preview::Line::Refused, L"op " + std::to_wstring (s.op) + L"  \u00B7  " + s.type,
										   s.why));
			}
		int restoreOps = 0;
		{
		long lastOp = -1;
		for (const Undo::Restore &x : r.restores)
			if (x.change.op != lastOp)
				{
				++restoreOps;
				lastOp = x.change.op;
				}
		}
		if (!r.restores.empty ())
			{
			lines.push_back (MakeLine (Preview::Line::Section, L"Restore - " + std::to_wstring (r.restores.size ())
									   + (r.restores.size () == 1 ? L" value on " : L" values on ")
									   + std::to_wstring (restoreOps) + (restoreOps == 1 ? L" operation" : L" operations")));
			long lastOp = -1;
			for (size_t k = 0; k < r.restores.size (); ++k)
				{
				const Undo::Restore &x = r.restores[k];
				if (x.change.op != lastOp)
					{
					Preview::Line o = MakeLine (Preview::Line::Op, L"op " + std::to_wstring (x.change.op) + L"  \u00B7  "
												+ x.type, comments[x.change.op]);
					o.box = Preview::Line::Ticked;
					o.tag = x.change.op;
					lines.push_back (o);
					lastOp = x.change.op;
					}
				Preview::Line c = MakeLine (Preview::Line::Change, x.change.name, std::wstring (),
											x.change.from, x.change.to);
				c.box = Preview::Line::Ticked;
				c.link = Plan::LinkOf (x.change.name);
				c.tag = static_cast<long> (k);
				if (changeLine < 0)
					changeLine = static_cast<int> (lines.size ());
				lines.push_back (c);
				}
			}
		if (!r.already.empty ())
			{
			lines.push_back (MakeLine (Preview::Line::Section, L"Already as before the load ("
									   + std::to_wstring (r.already.size ()) + L")"));
			for (const Undo::Entry &e : r.already)
				lines.push_back (MakeLine (Preview::Line::Note, L"op " + std::to_wstring (e.op) + L"  \u00B7  " + e.type,
										   e.column));
			}

		const std::wstring summary =
			L"Load of " + last.stamp + (last.files.empty () ? L"" : L" from " + last.files)
			+ L"  \u00B7  " + std::to_wstring (r.restores.size ()) + L" to restore"
			+ (r.skipped.empty () ? L"" : L"  \u00B7  " + std::to_wstring (r.skipped.size ()) + L" changed since")
			+ (r.already.empty () ? L"" : L"  \u00B7  " + std::to_wstring (r.already.size ()) + L" already back");

		Preview::Options options;
		options.caption = L"Parameter Table Tool - undo last load";
		options.verb = L"Restore";
		options.foot = L"Untick a value to keep it as it is now. Every value replaced goes to ParamTable.log "
					   L"first. Restored operations are marked for regeneration.";
		options.nothing = last.entries.empty () ? L"That load wrote nothing - there is nothing to restore."
												: L"Nothing to restore.";

		Scene scene;
		scene.name = L"undo";
		scene.fill = FitToLines;
		scene.callouts = [&lines, &summary, sinceLine, changeLine] (HWND dlg, const RECT &frame, int dpi)
			{
			std::vector<Callout> c;
			HWND list = ListOf (dlg);
			HWND restore = GetDlgItem (dlg, IDOK);
			const HFONT bold = reinterpret_cast<HFONT> (SendMessage (restore, WM_GETFONT, 0, 0));
			// 1: after the line naming the load.
			for (HWND s : Controls (dlg, [] (HWND x) { return ClassOf (x) == L"Static"; }))
				if (TextOf (s) == summary)
					{
					const RECT r = InPicture (s, frame);
					c.push_back (RightOf (1, r.left + TextW (s, summary), r.top + TextH (s) / 2, dpi));
					}
			// 2: after a value going back.
			if (list != nullptr && changeLine >= 0)
				{
				const RECT r = LineInPicture (list, changeLine, 3, frame);
				c.push_back (RightOf (2, r.left + 6 + TextW (list, lines[static_cast<size_t> (changeLine)].to, bold),
									  (r.top + r.bottom) / 2, dpi));
				}
			// 3: after the heading of what changed since.
			if (list != nullptr && sinceLine >= 0)
				{
				const RECT r = LineInPicture (list, sinceLine, 0, frame);
				c.push_back (RightOf (3, r.left + 8 + TextW (list, lines[static_cast<size_t> (sinceLine)].text, bold),
									  (r.top + r.bottom) / 2, dpi));
				}
			c.push_back (LeftOf (4, InPicture (restore, frame), dpi));
			return c;
			};
		return Make (scene, [&] { Preview::Show (L"Undo the last load of " + part, summary, lines, options); });
		}
	}

int wmain (int argc, wchar_t **argv)
	{
	// Drawn at the screen's own scale (as Mastercam and Excel draw), not blurred up to it.
	SetProcessDPIAware ();
	if (!AfxWinInit (GetModuleHandle (nullptr), nullptr, GetCommandLine (), 0))
		{
		std::puts ("dialog_shots: MFC did not start");
		return 1;
		}

	Args a;
	if (!ReadArgs (argc, argv, a))
		{
		std::puts ("usage: dialog_shots <workbook> <output folder> [--part p] [--edit op:column]... "
				   "[--stamp t] [--since op:column=value] [--find words] [--saved name] [--only dump|load|undo|batch]");
		return 1;
		}
	gOut = a.out;
	std::error_code ec;
	std::filesystem::create_directories (gOut, ec);

	Sheet sheet, dumped;
	std::wstring why;
	if (!Xlsx::ReadSheet (a.workbook, sheet.rows, sheet.rowNo, why))
		{
		std::printf ("dialog_shots: %ls\n", why.c_str ());
		return 1;
		}
	if (!Xlsx::ReadNamedSheet (a.workbook, L"Dumped", dumped.rows, dumped.rowNo, why))
		dumped.rows.clear ();

	// Each value the load writes: as dumped (the Dumped sheet - what the op holds)
	// and as on the sheet. One that the sheet does not change is a story that no
	// longer fits the workbook: said, and left out.
	std::vector<Edit> edits;
	const int cType = sheet.Col (L"type");
	for (const auto &e : a.edits)
		{
		Edit x;
		x.op = e.first;
		x.column = e.second;
		const size_t r = sheet.RowOf (x.op);
		x.col = sheet.Col (x.column);
		if (r == 0 || x.col < 0)
			{
			std::printf ("  note   op %ld %ls is not on the sheet - left out\n", x.op, x.column.c_str ());
			continue;
			}
		x.row = sheet.rowNo[r];
		x.type = sheet.Cell (r, cType);
		x.to = sheet.Cell (r, x.col);
		x.from = dumped.rows.empty () ? x.to : dumped.Cell (x.op, x.column);
		if (x.from == x.to)
			{
			std::printf ("  note   op %ld %ls is not changed on the sheet - left out\n", x.op, x.column.c_str ());
			continue;
			}
		edits.push_back (x);
		}
	const int cComment = sheet.Col (L"comment");
	for (size_t k = 1; k < sheet.rows.size (); ++k)
		{
		long long id = 0;
		if (Csv::ParseLong (sheet.Cell (k, 0), id))
			a.comments[static_cast<long> (id)] = sheet.Cell (k, cComment);
		}

	const auto want = [&a] (const wchar_t *name) { return a.only.empty () || a.only.count (name) != 0; };
	int failed = 0;
	if (want (L"dump"))
		failed += !DumpWindow (a, sheet);
	if (want (L"load"))
		failed += !LoadPreview (a, sheet, dumped.rows, edits);
	if (want (L"undo"))
		failed += !UndoWindow (a, edits);
	if (want (L"batch"))
		failed += !BatchWindow (a);
	return failed == 0 ? 0 : 1;
	}
