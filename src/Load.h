#pragma once

namespace Load
	{
	/// Read edited CSVs back and write their changes into the operations.
	int Run ();

	/// Put the old values of this part's most recent load back, from
	/// ParamTable.log - each one only where the load's value is still there.
	int UndoLast ();
	}
