//
// FileRules.h - where a dump goes and what it is called. No Mastercam SDK.
//
// A name is a PATTERN with tokens, filled in at dump time:
//
//   {part}   the part file's name, without extension
//   {date}   yyyymmdd
//   {time}   hhmmss
//   {scope}  "selected" or "all"
//   {ops}    how many operations
//
// Characters Windows does not allow in a file name become '_', the extension
// (".xlsx", or ".xlsm" with macros) is added, and an existing file is NEVER
// overwritten - " (2)", " (3)" ... is added instead.
//
#pragma once

#include <ctime>
#include <filesystem>
#include <string>

namespace FileRules
	{
	const wchar_t *const kDefaultPattern = L"{part}_lathe_params_{date}-{time}";

	/// The pattern filled in, made safe as a file name, with `ext` (".xlsx" / ".xlsm").
	std::wstring Name (const std::wstring &pattern, const std::wstring &part, std::time_t when,
					   bool selectedOnly, size_t ops, const std::wstring &ext = L".xlsx");

	/// `folder / name`, or the first "name (n).xlsx" that does not exist yet.
	std::filesystem::path Unique (const std::filesystem::path &folder, const std::wstring &name);
	}
