//
// Batch.h - "Parameter Table - dump a folder of parts": continuous improvement across
// the shop, not one part at a time. Every part of a folder (and its subfolders, if
// asked) is opened in turn - not regenerated, never saved - dumped exactly as "Lathe
// params - dump to Excel" dumps a whole part (Dump::Quiet: the same function, no
// window, no Excel), and closed. Then one SHOP WORKBOOK puts them side by side
// (Shop.h), and the part that was open is opened again.
//
// Mastercam has one part open at a time, so the part on screen must be SAVED first:
// with unsaved changes the batch does not start (IFileManager::IsFileDirty).
//
// What happens to every part goes to a batch log beside the shop workbook - with
// what the SDK said at each step (open, the file now current, unsaved changes before
// and after the dump), so a real run shows how Mastercam took it. A part that will
// not open or dump is passed over, and why is said in the log and on the Parts
// sheet. A Cancel stops after the part in hand; the parts done are kept.
//
#pragma once

namespace Batch
	{
	/// Ask (the batch window), dump every part, write the shop workbook, open the
	/// first part again, and say how it went.
	int Run ();
	}
