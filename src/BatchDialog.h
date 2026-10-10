//
// BatchDialog.h - the batch dump's two windows: what to dump and where (Show), and
// the one it runs under, part by part, whose Cancel stops it after the part in
// hand (Progress). No Mastercam SDK (Ui.h), as the dump window.
//
#pragma once

#include "Settings.h"

#include <memory>
#include <string>

namespace BatchDialog
	{
	/// Ask: the folder of parts (and its subfolders?), where the workbooks go and
	/// the options - with the warning that the open part is closed and opened again
	/// at the end. `openPart` is the part open now ("" = none saved), `ext` the part
	/// extension (".mcam"). `settings` come in as the defaults and go out as chosen.
	/// False = cancelled.
	bool Show (const std::wstring &openPart, const std::wstring &ext, Settings::Batch &settings);

	/// The window the batch runs under: which part, how far, and Cancel. While it is
	/// up Mastercam's own window takes no clicks - a click queued during a dump would
	/// otherwise land between two parts, in the middle of the batch.
	class Progress
		{
		public:
			explicit Progress (size_t parts);
			~Progress ();
			Progress (const Progress &) = delete;
			Progress &operator= (const Progress &) = delete;

			/// The part about to be dumped (`index` from 0) and how the others went;
			/// then the window answers its buttons. False once Cancel was pressed:
			/// stop before this part.
			bool Next (size_t index, const std::wstring &name, size_t done, size_t failed);

			/// What is done after the parts ("Opening Shaft.mcam again ...").
			void Finishing (const std::wstring &what);

		private:
			struct Impl;
			std::unique_ptr<Impl> m;
		};
	}
