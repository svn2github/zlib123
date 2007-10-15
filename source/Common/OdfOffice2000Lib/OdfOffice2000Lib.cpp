#include "stdafx.h"
#include "OdfOffice2000Lib.h"


void TraceError(HRESULT hr, LPCTSTR sTitle)
{
#if _DEBUG
	TCHAR sBuf[512];
	if (!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, hr, 0, sBuf, 512, NULL)) {
		StringCchPrintf(sBuf, 512, _T("Err 0x%08X"), hr);
	}
	OutputDebugString(sTitle);
	OutputDebugString(_T(" : "));
	OutputDebugString(sBuf);
	OutputDebugString(_T("\n"));
#endif
}