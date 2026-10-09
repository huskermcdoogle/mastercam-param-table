#include "stdafx.h"
#include "MastercamSdk.h"
#include "ToolPictures.h"
#include "TlServices_CH.h"
#include "TlMgr_CH.h"
#include "ILTool_CH.h"
#include "SetupSheet_CH.h"

#include <atlimage.h>
#include <filesystem>
#include <fstream>
#include <iterator>

namespace ToolPictures
	{
	std::wstring LatheInsert (long slot)
		{
		Cnc::Tool::TlMgr *mgr = Cnc::Tool::GetTlMgr ();
		Cnc::Tool::ILToolCPtr tool;
		if (mgr == nullptr || !mgr->Find (slot, tool) || !tool)
			return std::wstring ();
		const CString name = tool->GetInsertName ();
		return std::wstring (name.GetString ());
		}

	bool LatheTool (long slot, std::string &png, int &width, int &height, std::wstring &why)
		{
		png.clear ();
		width = height = 0;
		Cnc::Tool::TlMgr *mgr = Cnc::Tool::GetTlMgr ();
		if (mgr == nullptr)
			{
			why = L"no tool manager";
			return false;
			}
		Cnc::Tool::ILToolCPtr tool;
		if (!mgr->Find (slot, tool) || !tool)
			{
			why = L"not a lathe tool";
			return false;
			}

		std::wstring bmp, holder, insert;
		if (!WriteLatheToolImage (tool, true, bmp, false, holder, false, insert, true) || bmp.empty ())
			{
			why = L"Mastercam drew no picture";
			return false;
			}

		CImage img;
		if (FAILED (img.Load (bmp.c_str ())))
			{
			why = L"could not read " + bmp;
			return false;
			}
		width = img.GetWidth ();
		height = img.GetHeight ();

		wchar_t tmp[MAX_PATH] = L"";
		GetTempPathW (MAX_PATH, tmp);
		const std::filesystem::path out = std::filesystem::path (tmp)
										  / (L"ParamTable_tool_" + std::to_wstring (slot) + L".png");
		if (FAILED (img.Save (out.c_str (), Gdiplus::ImageFormatPNG)))
			{
			why = L"could not convert the picture";
			return false;
			}
		{
		std::ifstream in (out, std::ios::binary);
		png.assign ((std::istreambuf_iterator<char> (in)), std::istreambuf_iterator<char> ());
		}
		std::error_code ec;
		std::filesystem::remove (out, ec);
		if (png.empty ())
			{
			why = L"empty picture";
			return false;
			}
		return true;
		}
	}
