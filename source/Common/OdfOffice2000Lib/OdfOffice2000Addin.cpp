#include "stdafx.h"
#include "OdfOffice2000Lib.h"

OdfOffice2000Addin::OdfOffice2000Addin(HMODULE hModule, UINT nDefaultTitle) : m_importHandler(this, ID_IMPORT), m_exportHandler(this, ID_EXPORT)
{
	m_pImportButton = NULL;
	m_pExportButton = NULL;
	m_hModule = hModule;
	m_nDefaultTitle = nDefaultTitle;
	m_pPreviousFilter = NULL;
}


OdfOffice2000Addin::~OdfOffice2000Addin()
{
	if (m_pPreviousFilter != NULL) {
		UninstallMessageFilter();
	}
}




void OdfOffice2000Addin::ShowMessage(UINT idMsg, UINT idTitle)
{
	TCHAR sBuf[512];
	TCHAR sTitle[512];
	LoadString(idMsg, sBuf, 512);
	LoadString(idTitle, sTitle, 512);
	MessageBox(NULL, sBuf, sTitle, MB_OK |  MB_ICONERROR);
}

void OdfOffice2000Addin::ShowMessage(UINT idMsg)
{
	ShowMessage(idMsg, m_nDefaultTitle);
}


void OdfOffice2000Addin::LoadString(UINT uID, LPTSTR lpBuffer, int nBufferMax)
{
#if FALSE
	::LoadString(m_hModule, uID, lpBuffer, nBufferMax);
#else
	HRSRC hrsrc = FindResourceEx(m_hModule, RT_STRING, MAKEINTRESOURCE(uID / 16 + 1), m_wLanguage);
	if (hrsrc == NULL) {
		TraceError(GetLastError(), _T("FindResourceEx"));
		::LoadString(m_hModule, uID, lpBuffer, nBufferMax);
	} else {
		HGLOBAL hRes = LoadResource(m_hModule, hrsrc);
		LPBYTE ptr = (LPBYTE)LockResource(hRes);
		WORD len;
		LPTSTR psz;
		for (int i = 0; i <= (uID % 16); i++) {
			WORD *pwLen = (WORD*)ptr;
			len = *pwLen;
			psz = (LPTSTR)(ptr + 2);
			// Move to next string
			ptr += sizeof(WORD) + 2 * len;
		}
		_tcsncpy(lpBuffer, psz, nBufferMax-1);
		if (len < nBufferMax) {
			lpBuffer[len] = 0;
		} else {
			lpBuffer[nBufferMax - 1] = 0;
		}
	}

#endif
}


// Temporary filename generation
BOOL OdfOffice2000Addin::GenerateTempName(LPTSTR sBuffer, DWORD dwBufferLen, LPCTSTR sRoot, LPCTSTR sExtension)
{

	if (sRoot) {
		// sRoot given, the target pattern is <TEMPFOLDER>\<sRoot>(withoutExtension)<UniqueNumber>.<sExtension>
		DWORD dwLen = GetTempPath(dwBufferLen, sBuffer);
		if (dwLen == 0) return FALSE;
		if (sBuffer[dwLen-1] != L'\\') {
			sBuffer[dwLen++] = L'\\';
		}
		LPCTSTR sOldExtension = PathFindExtension(sRoot);
		int nRootLen = (int)(sOldExtension - sRoot);
		if (FAILED(StringCchCopyN(sBuffer + dwLen, dwBufferLen - dwLen, sRoot, nRootLen))) {
			return FALSE;
		} else {
			dwLen += nRootLen;
		}
		int number = 0;
		while (TRUE) {
			// Generate name
			DWORD dwNewLen = dwLen;
			if (number != 0) {
				if (FAILED(StringCchPrintf(sBuffer + dwLen, dwBufferLen - dwLen, _T("%d"), number))) {
					return FALSE;
				} else {
					dwNewLen += (int)wcslen(sBuffer + dwLen);
				}
			}
			// Add extension
			if (FAILED(StringCchPrintf(sBuffer + dwNewLen, dwBufferLen - dwNewLen, sExtension))) {
				return FALSE;
			}
			// Now the name is complete
			if (GetFileAttributes(sBuffer) == (DWORD)-1) {
				// File does not exist, ok
				return TRUE;
			}
			number++;
		}
	} else {
		// Not implemented yet
		return FALSE;
	}
}



HRESULT OdfOffice2000Addin::ConfigureMenus(_CommandBars *pCommandBars, UINT uExportLabel, UINT uImportLabel, Office2000::CommandBarControl **ppOpenButton)
{
	// Variables to cleanup before exit
	BSTR bstrCommandBarName = SysAllocString(L"File");
	CommandBar *pCommandBar = NULL;
	CommandBarControls *pControls = NULL;

	// Register menu items
	VARIANT vtCommandBarName;
	vtCommandBarName.vt = VT_BSTR;
	vtCommandBarName.bstrVal = bstrCommandBarName;


	HRESULT hr = pCommandBars->get_Item(vtCommandBarName, &pCommandBar);
	if (FAILED(hr)) {
		TraceError(hr, _T("CommandBars[\"File\"]"));
		goto Cleanup;
	}

#ifdef _DEBUG
	{
		TCHAR sTmp[128];
		StringCchPrintf(sTmp, 128, _T("pCommandBar = 0x%08X\n"), pCommandBar);
		OutputDebugString(sTmp);
	}
#endif

	hr = pCommandBar->get_Controls(&pControls);
	if (FAILED(hr)) {
		TraceError(hr, _T("get_Controls()"));
		goto Cleanup;
	}

	if (uExportLabel) {
		hr = CreateButton(pControls, uExportLabel, ID_EXPORT, &m_pExportButton, &m_exportHandler, &m_dwExportAdvise);
		if (FAILED(hr)) {
			goto Cleanup;
		}
	}

	if (uImportLabel) {
		hr = CreateButton(pControls, uImportLabel, ID_IMPORT, &m_pImportButton, &m_importHandler, &m_dwImportAdvise);
		if (FAILED(hr)) {
			goto Cleanup;
		}
	}

	// "Dirty" hack: return Open button (used only in Excel...)
	if (ppOpenButton) {
		// Locate and return OpenButton
		VARIANT vtType;
		vtType.vt = VT_I4;
		vtType.lVal = msoControlButton;
		VARIANT vtId;
		vtId.vt = VT_I4;
		vtId.lVal = 23;
		pCommandBar->FindControl(vtType, vtId, vtMissing, vtMissing, vtMissing, ppOpenButton);
	}

Cleanup:
	if (pCommandBar) pCommandBar->Release();
	if (pControls) pControls->Release();
	if (bstrCommandBarName) SysFreeString(bstrCommandBarName);

	return hr;
}

HRESULT OdfOffice2000Addin::CreateButton(CommandBarControls *pControls, UINT idResource, int idCommand, CommandBarControl **ppButton, CClickHandler *pClickHandler, DWORD *pdwAdvise)
{

	// Variables to cleanup before exit
	BSTR bstrCaption = NULL;

	// Register Import label and command
	TCHAR sCaption[256];
#if TRUE
	LoadString(idResource, sCaption, 256);
#else
	FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, m_hModule, idResource, m_wLanguage, sCaption, 256, NULL);
#endif
	bstrCaption = SysAllocString(sCaption);

	VARIANT vtCaption;
	vtCaption.vt = VT_BSTR;
	vtCaption.bstrVal = bstrCaption;

	HRESULT hr = pControls->get_Item(vtCaption, ppButton);
	if (SUCCEEDED(hr)) {
#ifdef _DEBUG
		OutputDebugString(_T("Office2000Addin:button already present\n"));
#endif
	} else {
#ifdef _DEBUG
		OutputDebugString(_T("Office2000Addin:button needs to be created\n"));
#endif
		VARIANT vtType;
		vtType.vt = VT_I4;
		vtType.intVal = 1;

		VARIANT vtId;
		vtId.vt = VT_I4;
		vtId.intVal = 1;

		VARIANT vtBefore;
		vtBefore.vt = VT_I4;
		vtBefore.intVal = 3;

		VARIANT vtTemporary;
		vtTemporary.vt = VT_BOOL;
		vtTemporary.boolVal = VARIANT_TRUE;

		hr = pControls->Add(vtType, vtId, vtMissing, vtBefore, vtTemporary, ppButton);
		if (FAILED(hr)) {
#ifdef _DEBUG
			TraceError(hr, _T("pControls->Add"));
#endif
			goto Cleanup;
		}
		(*ppButton)->put_Caption(bstrCaption);
		(*ppButton)->put_Tag(bstrCaption);
	}

	hr = AtlAdvise(*ppButton, pClickHandler, __uuidof(_CommandBarButtonEvents), pdwAdvise);
	if (FAILED(hr)) {
		TraceError(hr, _T("AtlAdvise"));
		return E_FAIL;
	}


Cleanup:
	if (bstrCaption) SysFreeString(bstrCaption);
	return hr;
}

HRESULT OdfOffice2000Addin::CleanMenus()
{
	if (m_pImportButton) {
		AtlUnadvise(m_pImportButton, __uuidof(_CommandBarButtonEvents), m_dwImportAdvise);
		VARIANT vtTemporary;
		vtTemporary.vt = VT_BOOL;
		vtTemporary.boolVal = VARIANT_TRUE;
		OutputDebugString(_T("Appel  de m_pImportButton->Delete(vtTemporary)\n"));
		TraceError(m_pImportButton->Delete(vtTemporary), _T("Delete"));
		OutputDebugString(_T("Retour de m_pImportButton->Delete(vtTemporary)\n"));
		m_pImportButton->Release();
		m_pImportButton = NULL;
	}

	if (m_pExportButton) {
		AtlUnadvise(m_pExportButton, __uuidof(_CommandBarButtonEvents), m_dwExportAdvise);
		VARIANT vtTemporary;
		vtTemporary.vt = VT_BOOL;
		vtTemporary.boolVal = VARIANT_TRUE;
		OutputDebugString(_T("Appel  de m_pExportButton->Delete(vtTemporary)\n"));
		TraceError(m_pExportButton->Delete(vtTemporary), _T("Delete"));
		OutputDebugString(_T("Retour de m_pExportButton->Delete(vtTemporary)\n"));
		m_pExportButton->Release();
		m_pExportButton = NULL;
	}

	return S_OK;
}



DWORD OdfOffice2000Addin::HandleInComingCall(DWORD dwCallType, HTASK htaskCaller, DWORD dwTickCount, LPINTERFACEINFO lpInterfaceInfo)
{
	if (m_pPreviousFilter) {
		return m_pPreviousFilter->HandleInComingCall(dwCallType, htaskCaller, dwTickCount, lpInterfaceInfo);
	} else {
		return SERVERCALL_ISHANDLED;
	}
}

DWORD OdfOffice2000Addin::RetryRejectedCall(HTASK htaskCallee, DWORD dwTickCount, DWORD dwRejectType)
{
	if (m_pPreviousFilter) {
		return RetryRejectedCall(htaskCallee, dwTickCount, dwRejectType);
	} else {
		return (DWORD)-1;
	}
}

DWORD OdfOffice2000Addin::MessagePending(HTASK htaskCallee, DWORD dwTickCount, DWORD dwPendingType)
{
	MSG msg;
	while (::PeekMessage(&msg, NULL, WM_PAINT, WM_PAINT, PM_REMOVE|PM_NOYIELD))
	{
		DispatchMessage(&msg);
	}
	if (m_pPreviousFilter) {
		return m_pPreviousFilter->MessagePending(htaskCallee, dwTickCount, dwPendingType);
	} else {
		return PENDINGMSG_WAITNOPROCESS;
	}
}

void OdfOffice2000Addin::InstallMessageFilter()
{
	HRESULT hr = CoRegisterMessageFilter((IMessageFilter*)this, &m_pPreviousFilter);
	if (FAILED(hr)) {
		TraceError(hr, _T("CoRegisterMessageFilter(register)"));
	} else {
#ifdef _DEBUG
		OutputDebugString(_T("MessageFilter installed\n"));
#endif
	}
}

void OdfOffice2000Addin::UninstallMessageFilter()
{
	IMessageFilter *pPrevious = NULL;
	HRESULT hr = CoRegisterMessageFilter(m_pPreviousFilter, &pPrevious);
	if (FAILED(hr)) {
		TraceError(hr, _T("CoRegisterMessageFilter(unregister)"));
	} else {
#ifdef _DEBUG
		OutputDebugString(_T("MessageFilter removed\n"));
#endif
	}
	if (pPrevious) pPrevious->Release();
	if (m_pPreviousFilter) {
		m_pPreviousFilter->Release();
		m_pPreviousFilter = NULL;
	}
}

