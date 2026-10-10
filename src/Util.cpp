#include "stdafx.h"
#include "MastercamSdk.h"
#include "Util.h"
#include "Ui.h"
#include "Csv.h"

#include <ctime>
#include <fstream>

namespace Ui
	{
	CWnd *Host ()
		{
		return CWnd::FromHandle (get_MainFrame ()->GetSafeHwnd ());
		}
	}

namespace Util
	{
	std::filesystem::path PartFile ()
		{
		try
			{
			auto fm = Mastercam::CHookAPI::IO::GetFileManager ();
			if (fm)
				return std::filesystem::path (fm->GetCurrentFileName ());
			}
		catch (...)
			{
			}
		return {};
		}

	int Say (const std::wstring &text, unsigned flags)
		{
		return ::MessageBoxW (get_MainFrame ()->GetSafeHwnd (), text.c_str (),
							  L"Parameter Table Tool", flags);
		}

	void Log (const std::filesystem::path &partFile, const std::wstring &line)
		{
		if (partFile.empty ())
			return;

		std::filesystem::path log = partFile.parent_path () / L"ParamTable.log";

		std::time_t now = std::time (nullptr);
		std::tm parts {};
		localtime_s (&parts, &now);
		wchar_t stamp[32];
		wcsftime (stamp, _countof (stamp), L"%Y-%m-%d %H:%M:%S", &parts);

		const std::wstring text = std::wstring (stamp) + L"  " + line + L"\r\n";
		const std::string bytes = Csv::ToUtf8Bom (text);

		// A BOM only at the top of a NEW file - appended after that it would be
		// a stray character in the middle of the log.
		const bool fresh = !std::filesystem::exists (log);
		std::ofstream out (log, std::ios::binary | std::ios::app);
		if (!out)
			return;
		out.write (fresh ? bytes.data () : bytes.data () + 3,
				   static_cast<std::streamsize> (fresh ? bytes.size ()
													   : bytes.size () - 3));
		}
	}
