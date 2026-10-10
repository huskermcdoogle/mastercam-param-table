#include "stdafx.h"
#include "MastercamSdk.h"
#include "ToolPictures.h"
#include "TlServices_CH.h"
#include "TlMgr_CH.h"
#include "ILTool_CH.h"
#include "TlToolLathe_CH.h"
#include "TlHolderLathe_CH.h"
#include "TlInsert_CH.h"
#include "TlToolGrade_CH.h"
#include "SetupSheet_CH.h"
#include "TlToolMill_CH.h"
#include "ITlBitmapFactory_CH.h"

#include <atlimage.h>
#include <filesystem>
#include <fstream>
#include <cwctype>
#include <iterator>

namespace ToolPictures
	{
	/// A picture as PNG bytes - through a temp file, which is what CImage saves
	/// to dependably - and its size in pixels.
	bool ToPng (CImage &img, long slot, std::string &png, int &width, int &height, std::wstring &why);

	InsertInfo LatheInsertInfo (long slot)
		{
		InsertInfo r;
		Cnc::Tool::TlMgr *mgr = Cnc::Tool::GetTlMgr ();
		Cnc::Tool::TlToolLathePtr tool;
		if (mgr == nullptr || !mgr->Find (slot, tool) || !tool)
			return r;
		const auto ins = tool->GetMainInsert ();
		if (!ins)
			return r;
		r.ok = true;
		r.shape = static_cast<wchar_t> (std::towupper (ins->GetAnsiShapeCode ()));
		r.ic = ins->GetICDiameter ();
		r.radius = ins->GetCornerRadius ();
		r.thickness = ins->GetThickness ();
		r.width = ins->GetWidth ();
		r.length = ins->GetLength ();
		r.custom = ins->GetIsCustom ();
		if (const auto g = ins->GetGrade ())
			r.grade = g->GetName ().GetString ();
		return r;
		}

	HolderInfo LatheHolderInfo (long slot)
		{
		HolderInfo r;
		// Anything the tool manager throws is a tool with no holder data, never a
		// dump that stops.
		try
			{
			Cnc::Tool::TlMgr *mgr = Cnc::Tool::GetTlMgr ();
			Cnc::Tool::TlToolLathePtr tool;
			if (mgr == nullptr || !mgr->Find (slot, tool) || !tool)
				return r;
			Cnc::Tool::ILToolCPtr named;
			if (mgr->Find (slot, named) && named)
				r.name = named->GetHolderName ().GetString ();
			const auto holder = tool->GetMainHolder ();
			if (!holder)
				return r;
			r.ok = true;
			r.style = static_cast<wchar_t> (holder->GetAnsiShapeCode ());
			r.type = static_cast<int> (holder->GetType ());
			r.insertShapes = holder->GetInsertShapes ().GetString ();
			r.sideAngle = holder->GetSideCuttingEdgeAngle ();
			r.endAngle = holder->GetEndCuttingEdgeAngle ();
			}
		catch (...)
			{
			r.ok = false;
			}
		return r;
		}

	std::wstring LatheMfgCode (long slot)
		{
		Cnc::Tool::TlMgr *mgr = Cnc::Tool::GetTlMgr ();
		Cnc::Tool::ILToolCPtr tool;
		if (mgr == nullptr || !mgr->Find (slot, tool) || !tool)
			return std::wstring ();
		return std::wstring (tool->GetMfgCode ().GetString ());
		}

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
		return ToPng (img, slot, png, width, height, why);
		}

	bool MillTool (long slot, std::string &png, int &width, int &height, std::wstring &why)
		{
		png.clear ();
		width = height = 0;
		// Anything the tool manager throws is a tool with no picture, never a
		// dump that stops.
		try
			{
			Cnc::Tool::TlMgr *mgr = Cnc::Tool::GetTlMgr ();
			if (mgr == nullptr)
				{
				why = L"no tool manager";
				return false;
				}
			// Through the MILL slot table: a lathe tool can have the same slot number.
			std::shared_ptr<const Cnc::Tool::TlToolMill> tool;
			if (!mgr->Find (mgr->IDBySlot (slot), tool) || !tool)
				{
				why = L"not a mill tool";
				return false;
				}

			// THE TOOL MANAGER'S OWN BITMAP FACTORY, not the setup sheet's
			// WriteMillToolImage: setup-sheet code reports on operations, and on a
			// part with two tools sharing a number it asks about them - once per
			// call. The factory only draws the tool.
			const Cnc::Tool::ITlBitmapFactoryPtr factory = Cnc::Tool::CreateITlBitmapFactory ();
			if (!factory)
				{
				why = L"no tool bitmap factory";
				return false;
				}
			Cnc::Tool::ITlBitmapFactory::ToolParams params;
			params.size = Cnc::Tool::ITlBitmapFactory::Size::SetupSheet;
			params.backgroundColor = RGB (255, 255, 255);
			const std::unique_ptr<CBitmap> bmp = factory->Create (*tool, params);
			BITMAP bm = {};
			if (!bmp || bmp->GetSafeHandle () == nullptr || !bmp->GetBitmap (&bm)
				|| bm.bmWidth <= 0 || bm.bmHeight <= 0)
				{
				why = L"Mastercam drew no picture";
				return false;
				}

			// Copied onto a white 24-bit image: the factory's bitmap may carry an
			// alpha channel of zeros, which a PNG would show as nothing at all.
			CImage img;
			if (!img.Create (bm.bmWidth, bm.bmHeight, 24))
				{
				why = L"could not make the picture";
				return false;
				}
			bool copied = false;
			const HDC dst = img.GetDC ();
			const HDC src = CreateCompatibleDC (dst);
			if (src != nullptr)
				{
				const HGDIOBJ old = SelectObject (src, bmp->GetSafeHandle ());
				const RECT all = { 0, 0, bm.bmWidth, bm.bmHeight };
				FillRect (dst, &all, static_cast<HBRUSH> (GetStockObject (WHITE_BRUSH)));
				copied = BitBlt (dst, 0, 0, bm.bmWidth, bm.bmHeight, src, 0, 0, SRCCOPY) != FALSE;
				SelectObject (src, old);
				DeleteDC (src);
				}
			img.ReleaseDC ();
			if (!copied)
				{
				why = L"could not copy the picture";
				return false;
				}
			return ToPng (img, slot, png, width, height, why);
			}
		catch (...)
			{
			png.clear ();
			width = height = 0;
			why = L"the tool manager failed drawing it";
			return false;
			}
		}

	bool ToPng (CImage &img, long slot, std::string &png, int &width, int &height, std::wstring &why)
		{
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
