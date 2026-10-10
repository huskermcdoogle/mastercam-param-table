#include "stdafx.h"
#include "MastercamSdk.h"
#include "Batch.h"
#include "BatchDialog.h"
#include "Csv.h"
#include "Dump.h"
#include "FileRules.h"
#include "Settings.h"
#include "Shop.h"
#include "Util.h"

#include <shellapi.h>

#include <ctime>
#include <filesystem>
#include <fstream>

namespace
	{
	using Mastercam::CHookAPI::IO::IFileManagerPtr;

	/// A line of the batch log, time-stamped (UTF-8; a BOM only at the top of a new
	/// file, as ParamTable.log).
	void Note (const std::filesystem::path &log, const std::wstring &line)
		{
		std::time_t now = std::time (nullptr);
		std::tm parts {};
		localtime_s (&parts, &now);
		wchar_t stamp[32];
		wcsftime (stamp, _countof (stamp), L"%Y-%m-%d %H:%M:%S", &parts);
		const std::string bytes = Csv::ToUtf8Bom (std::wstring (stamp) + L"  " + line + L"\r\n");
		std::error_code ec;
		const bool fresh = !std::filesystem::exists (log, ec);
		std::ofstream out (log, std::ios::binary | std::ios::app);
		if (!out)
			return;
		out.write (fresh ? bytes.data () : bytes.data () + 3,
				   static_cast<std::streamsize> (fresh ? bytes.size () : bytes.size () - 3));
		}

	std::wstring YesNo (bool b)
		{
		return b ? L"yes" : L"no";
		}

	/// The same file, however Mastercam spells its name.
	bool SameFile (const std::wstring &a, const std::filesystem::path &b)
		{
		std::error_code ec;
		if (!a.empty () && std::filesystem::equivalent (std::filesystem::path (a), b, ec) && !ec)
			return true;
		return _wcsicmp (a.c_str (), b.c_str ()) == 0;
		}

	/// One part: opened (without asking to save the part before it - the batch
	/// saves nothing, so changes a dump may leave in a part are dropped with it),
	/// checked to be the part now open, and dumped whole. `p.ok`, or `p.why`.
	void DumpOne (const IFileManagerPtr &fm, const std::filesystem::path &file, const Settings::Dump &settings,
				  Shop::Part &p, const std::filesystem::path &log)
		{
		const bool opened = fm->Open (false, file.wstring ());
		const std::wstring now = fm->GetCurrentFileName ();
		Note (log, L"  open (IFileManager::Open, no prompt): " + YesNo (opened) + L" | current file now: "
				   + (now.empty () ? L"(none)" : now) + L" | unsaved changes: " + YesNo (fm->IsFileDirty ()));
		if (!opened || !SameFile (now, file))
			{
			p.why = opened ? L"Mastercam opened another file (" + now + L")" : L"Mastercam could not open it";
			Note (log, L"  NOT DUMPED: " + p.why);
			return;
			}
		Util::Log (file, L"batch dump: opened by \"dump a folder of parts\" - not regenerated, not saved");

		Dump::Outcome out;
		bool ok = false;
		try
			{
			ok = Dump::Quiet (settings, out);
			}
		catch (...)
			{
			ok = false;
			out.why = L"the dump failed unexpectedly (its ParamTable.log says how far it got)";
			}
		Note (log, L"  after the dump: unsaved changes " + YesNo (fm->IsFileDirty ()) + L" (never saved)");
		p.skipped = out.skipped;
		p.material = out.material;
		if (!ok)
			{
			p.why = out.why;
			Note (log, L"  NOT DUMPED: " + p.why);
			return;
			}
		p.ok = true;
		p.workbook = out.file.filename ().wstring ();
		Shop::Read (out.sheet, p);
		if (out.ops != p.ops.size ())
			p.why = std::to_wstring (out.ops) + L" ops dumped, " + std::to_wstring (p.ops.size ()) + L" read back";
		Note (log, L"  dumped " + std::to_wstring (out.ops) + L" op(s) to " + out.file.wstring ()
				   + L" | material: " + (p.material.empty () ? L"(none set)" : p.material)
				   + (p.skipped.empty () ? L"" : L" | not read yet: " + p.skipped)
				   + (p.regen > 0 ? L" | " + std::to_wstring (p.regen) + L" op(s) need regenerating" : L""));
		}

	/// The part open before the batch, open again - or, when none was saved, an
	/// empty part (no part of the batch is left on screen to be edited by mistake).
	bool Reopen (const IFileManagerPtr &fm, const std::wstring &open, const std::filesystem::path &log)
		{
		if (open.empty ())
			{
			const bool ok = fm->New (false);
			Note (log, L"no part was open before: an empty part (IFileManager::New, no prompt): " + YesNo (ok));
			return ok;
			}
		const bool ok = fm->Open (false, open);
		const std::wstring now = fm->GetCurrentFileName ();
		Note (log, L"the part open before, again: " + open + L" | open: " + YesNo (ok) + L" | current file now: "
				   + (now.empty () ? L"(none)" : now));
		return ok && SameFile (now, open);
		}
	}

namespace Batch
	{
	int Run ()
		{
		const IFileManagerPtr fm = Mastercam::CHookAPI::IO::GetFileManager ();
		if (!fm)
			{
			Util::Say (L"Mastercam's file manager did not answer, so no part can be opened. Nothing was done.", MB_ICONERROR);
			return 0;
			}

		// ---- THE PART ON SCREEN: saved, or nothing starts. Mastercam has one part
		// open at a time, so it is closed while the batch runs and opened again at
		// the end - unsaved changes would be lost.
		// (A name with no file behind it - a part never saved - counts as none.)
		const std::wstring current = fm->GetCurrentFileName ();
		std::error_code ec0;
		const std::wstring open = !current.empty () && std::filesystem::is_regular_file (current, ec0) ? current : std::wstring ();
		if (fm->IsFileDirty ())
			{
			Util::Say (L"Save the part first.\r\n\r\n"
					   + (open.empty () ? std::wstring (L"The part on screen has never been saved")
										: std::filesystem::path (open).filename ().wstring () + L" has changes that are not saved")
					   + L". Dumping a folder of parts closes it - Mastercam has one part open at a time - and opens "
						 L"it again at the end, so the changes would be lost.\r\n\r\n"
						 L"Save it (or close it), then run \"Parameter Table - dump a folder of parts\" again.",
					   MB_ICONWARNING);
			return 0;
			}
		std::wstring ext = fm->GetPartExtension ();
		if (ext.empty ())
			ext = L".mcam";
		if (ext[0] != L'.')
			ext = L"." + ext;

		Settings::Batch settings = Settings::LoadBatch ();
		if (!BatchDialog::Show (open, ext, settings))
			return 0;
		Settings::SaveBatch (settings);

		const std::filesystem::path root (settings.folder);
		bool more = false;
		const std::vector<std::filesystem::path> files = Shop::FindParts (root, settings.subfolders, ext, 0, more);
		if (files.empty ())
			{
			Util::Say (L"There are no " + ext + L" parts in\r\n\r\n  " + root.wstring () + L"\r\n\r\nNothing was done.",
					   MB_ICONWARNING);
			return 0;
			}

		// ---- Where the shop workbook goes - with the workbooks when they go to one
		// folder, else in the folder of parts - and its log beside it.
		const std::filesystem::path shopFolder = settings.outFolder.empty () ? root : std::filesystem::path (settings.outFolder);
		const std::time_t started = std::time (nullptr);
		std::filesystem::path shopFile = FileRules::Unique (shopFolder, Shop::FileName (started));
		const std::filesystem::path log = shopFile.parent_path () / (shopFile.stem ().wstring () + L".log");
		Note (log, L"batch dump: " + std::to_wstring (files.size ()) + L" part(s) in " + root.wstring ()
				   + (settings.subfolders ? L" and its subfolders" : L""));
		Note (log, L"  the part open now (GetCurrentFileName): " + (current.empty () ? std::wstring (L"(none)") : current)
				   + (open.empty () ? L" - not a saved part" : L"")
				   + L" | unsaved changes (IFileManager::IsFileDirty): no | part extension (GetPartExtension): "
				   + fm->GetPartExtension ());
		Note (log, L"  workbooks: " + (settings.outFolder.empty () ? std::wstring (L"beside each part") : settings.outFolder)
				   + L" | tool pictures " + YesNo (settings.pictures) + L" | macros " + YesNo (settings.macros)
				   + L" | stock simulation " + YesNo (settings.stockSim));

		// Each part dumped as a normal dump of it would be - its file name pattern too.
		Settings::Dump dump = Settings::LoadDump ();
		dump.folder = settings.outFolder;
		dump.openExcel = false;
		dump.pictures = settings.pictures;
		dump.macros = settings.macros;
		dump.stockSim = settings.stockSim;

		std::vector<Shop::Part> parts;
		size_t done = 0, failed = 0;
		bool stopped = false, written = false, reopened = false;
		{
		BatchDialog::Progress progress (files.size ());
		for (size_t i = 0; i < files.size (); ++i)
			{
			Shop::Part p;
			p.name = files[i].stem ().wstring ();
			p.file = files[i].lexically_relative (root).wstring ();
			if (p.file.empty ())
				p.file = files[i].wstring ();
			if (!progress.Next (i, p.file, done, failed))
				{
				stopped = true;
				Note (log, L"stopped (Cancel) before part " + std::to_wstring (i + 1) + L" of " + std::to_wstring (files.size ()));
				break;
				}
			Note (log, L"part " + std::to_wstring (i + 1) + L" of " + std::to_wstring (files.size ()) + L": " + files[i].wstring ());
			try
				{
				DumpOne (fm, files[i], dump, p, log);
				}
			catch (...)
				{
				p.ok = false;
				p.why = L"failed unexpectedly";
				Note (log, L"  NOT DUMPED: " + p.why);
				}
			if (p.ok)
				++done;
			else
				++failed;
			parts.push_back (std::move (p));
			}
		// The parts a Cancel left: listed, so the sheet says what is missing.
		for (size_t i = parts.size (); i < files.size (); ++i)
			{
			Shop::Part p;
			p.name = files[i].stem ().wstring ();
			p.file = files[i].lexically_relative (root).wstring ();
			p.why = L"not reached - stopped (Cancel)";
			parts.push_back (std::move (p));
			}

		// ---- The shop workbook: every part side by side.
		if (std::error_code ec; std::filesystem::exists (shopFile, ec))
			shopFile = FileRules::Unique (shopFolder, Shop::FileName (started));
		progress.Finishing (L"Writing " + shopFile.filename ().wstring ());
		wchar_t when[32] = L"";
		{
		std::tm local = {};
		if (localtime_s (&local, &started) == 0)
			std::wcsftime (when, 32, L"%Y-%m-%d %H:%M", &local);
		}
		try
			{
			written = Xlsx::WriteBook (shopFile, Shop::Tables (parts, L"Shop compare - " + root.wstring () + L" - " + when));
			}
		catch (...)
			{
			written = false;
			}
		Note (log, (written ? L"shop workbook: " : L"shop workbook: COULD NOT WRITE ") + shopFile.wstring ());

		// ---- The part that was open, open again.
		progress.Finishing (open.empty () ? std::wstring (L"Leaving an empty part open")
										  : L"Opening " + std::filesystem::path (open).filename ().wstring () + L" again");
		try
			{
			reopened = Reopen (fm, open, log);
			}
		catch (...)
			{
			reopened = false;
			Note (log, L"opening the first part again failed unexpectedly");
			}
		}
		Note (log, L"done: " + std::to_wstring (done) + L" dumped, " + std::to_wstring (failed) + L" not"
				   + (stopped ? L", stopped before the rest" : L""));

		// ---- Say how it went.
		std::wstring text = L"Dumped " + std::to_wstring (done) + L" of " + std::to_wstring (files.size ())
							+ (files.size () == 1 ? L" part." : L" parts.");
		if (failed > 0)
			{
			text += L"\r\n\r\nNot dumped:";
			size_t listed = 0;
			for (const Shop::Part &p : parts)
				if (!p.ok && p.why.rfind (L"not reached", 0) != 0 && listed++ < 8)
					text += L"\r\n  " + p.file + L" - " + p.why;
			if (failed > 8)
				text += L"\r\n  and " + std::to_wstring (failed - 8) + L" more - see the log";
			}
		if (stopped)
			text += L"\r\n\r\nStopped (Cancel) - " + std::to_wstring (files.size () - done - failed)
					+ L" part(s) not reached.";
		if (written)
			text += L"\r\n\r\nEvery part side by side:\r\n  " + shopFile.filename ().wstring () + L"\r\nin\r\n  "
					+ shopFile.parent_path ().wstring ();
		else
			text += L"\r\n\r\nThe shop workbook could not be written to " + shopFile.parent_path ().wstring ()
					+ L" - check the folder can be written to. Each part's own workbook is beside it.";
		text += L"\r\n\r\nWhat happened to each part: " + log.filename ().wstring () + L", in the same folder.";
		if (!open.empty ())
			text += reopened ? L"\r\n\r\n" + std::filesystem::path (open).filename ().wstring () + L" is open again."
							 : L"\r\n\r\nCould not open " + open + L" again - open it with File > Open.";
		if (!written)
			{
			Util::Say (text, MB_ICONWARNING);
			return 0;
			}
		if (Util::Say (text + L"\r\n\r\nOpen the shop workbook in Excel now?", MB_YESNO | MB_ICONINFORMATION) == IDYES)
			ShellExecuteW (nullptr, L"open", shopFile.c_str (), nullptr, nullptr, SW_SHOWNORMAL);
		return 0;
		}
	}
