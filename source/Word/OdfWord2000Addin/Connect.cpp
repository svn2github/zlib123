// Connect.cpp : Implementation of CConnect
#include "stdafx.h"
#include "AddIn.h"
#include "Connect.h"
#include "Resource.h"

static LPCTSTR LangOptionKey = _T("Software\\Clever Age\\Odf Add-in for Word");
static LPCTSTR LangOptionValue = _T("Language");

UINT_PTR CALLBACK MyHookProc(HWND hDlg, UINT nMsg, WPARAM wParam, LPARAM lParam)
{
	switch (nMsg) {
	case WM_INITDIALOG:
		{
			HWND hDialog = GetParent(hDlg);
			HWND hParent = GetParent(hDialog);
			RECT rcDialog, rcParent;
			GetWindowRect(hDialog, &rcDialog);
			GetWindowRect(hParent, &rcParent);
			SetWindowPos(hDialog,
						 NULL,
						 (rcParent.left + rcParent.right) / 2 - (rcDialog.right - rcDialog.left) / 2,
						 (rcParent.top + rcParent.bottom) / 2 - (rcDialog.bottom - rcDialog.top) / 2,
						 0,0,SWP_NOSIZE | SWP_NOZORDER);
		}
	default:
		return 0;
	}
}

CConnect::CConnect() : OdfOffice2000Addin((HMODULE)_AtlModule.GetResourceInstance(), IDS_ODTTITLE)
{
	m_wLanguage = LANGIDFROMLCID(GetThreadLocale());
	_Word12SaveFormat = -1;


	// Implementation fields
	m_pApplication = NULL;
	m_pDocuments = NULL;
}

CConnect::~CConnect()
{
}




// Get Word 2007 Format number
static LPCTSTR Word12Class = _T("Word12");

HRESULT CConnect::GetWord2007SaveFormat()
{
	// Variables to cleanup before exit
	FileConverters *pConverters = NULL;
	FileConverter *pConverter = NULL;
	IUnknown *pUnkEnum = NULL;
	IEnumVARIANT *pEnumConverters = NULL;
	VARIANT vtConverter;
	BSTR bstrClassName = NULL;

	VariantInit(&vtConverter);

	HRESULT hr = m_pApplication->get_FileConverters(&pConverters );
	if (FAILED(hr)) {
		TraceError(hr, _T("get_FileConverters()"));
		goto Cleanup;
	}

	hr = pConverters->get__NewEnum(&pUnkEnum);
	if (FAILED(hr)) {
		TraceError(hr, _T("get__NewEnum()"));
		goto Cleanup;
	}


	hr = pUnkEnum->QueryInterface(&pEnumConverters);
	if (FAILED(hr)) {
		TraceError(hr, _T("QueryInterface(&pEnumConverters)"));
		goto Cleanup;
	}

	hr = pEnumConverters->Reset();

	ULONG cFetched;
	while (SUCCEEDED(pEnumConverters->Next(1, &vtConverter, &cFetched)) && (cFetched == 1)) {
		if (vtConverter.vt == VT_DISPATCH) {
			hr = vtConverter.pdispVal->QueryInterface(&pConverter);
			if (FAILED(hr)) {
				TraceError(hr, _T("Unable to get item\n"));
			} else {
				hr = pConverter->get_ClassName(&bstrClassName);
				if (SUCCEEDED(hr)) {
					if (!_wcsicmp(bstrClassName, Word12Class)) {
						// Found it
						hr = pConverter->get_SaveFormat(&_Word12SaveFormat);
#ifdef _DEBUG
						TCHAR sTmp[256];
						StringCchPrintf(sTmp, 256, _T("Word12SaveFormat = %d\n"), _Word12SaveFormat);
						OutputDebugString(sTmp);
#endif
						goto Cleanup;
					}
					SysFreeString(bstrClassName);
					bstrClassName = NULL;
				} else {
					TraceError(hr, _T("Unable to get Classname\n"));
				}
				pConverter->Release();
				pConverter = NULL;
			}
		} else {
			OutputDebugString(_T("Not a VT_UNKNOWN\n"));
		}
		VariantClear(&vtConverter);
	}
#ifdef _DEBUG
	OutputDebugString(_T("Unable to find Word12SaveFormat\n"));
#endif

Cleanup:
	if (pUnkEnum) pUnkEnum->Release();
	if (pEnumConverters) pEnumConverters->Release();
	if (pConverter) pConverter->Release();
	if (pConverters) pConverters->Release();
	if (bstrClassName) SysFreeString(bstrClassName);
	return hr;
}


// When run, the Add-in wizard prepared the registry for the Add-in.
// At a later time, if the Add-in becomes unavailable for reasons such as:
//   1) You moved this project to a computer other than which is was originally created on.
//   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
//   3) Registry corruption.
// you will need to re-register the Add-in by building the OdfWord2000AddinSetup project, 
// right click the project in the Solution Explorer, then choose install.


// CConnect
STDMETHODIMP CConnect::OnConnection(IDispatch *pApplication, AddInDesignerObjects::ext_ConnectMode /*ConnectMode*/, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/ )
{
	HRESULT hr = pApplication->QueryInterface(&m_pApplication);
	if (FAILED(hr)) {
		TraceError(hr, _T("QueryInterface(_Application)"));
		ShowMessage(IDS_UNABLE_TO_CONNECT);
		return hr;
	}

	hr = m_pApplication->get_Documents(&m_pDocuments);
	if (FAILED(hr)) {
		TraceError(hr, _T("get_Documents()"));
		ShowMessage(IDS_UNABLE_TO_CONNECT);
		return hr;
	}

	// Now get language to use
	HKEY hKey;
	int nRet = RegOpenKeyEx(HKEY_CURRENT_USER, LangOptionKey, 0, GENERIC_READ, &hKey);
	if (nRet != ERROR_SUCCESS) {
		// Second try
		nRet = RegOpenKeyEx(HKEY_LOCAL_MACHINE, LangOptionKey, 0, GENERIC_READ, &hKey);
		if (nRet != ERROR_SUCCESS) {
			// Exit but don't consider it as an error
#ifdef _DEBUG
			OutputDebugString(_T("No Addin language option\n"));
#endif
			return hr;
		}
	}
	TCHAR sValue[32];
	DWORD dwSize = sizeof(sValue);
	nRet = RegQueryValueEx(hKey, LangOptionValue, NULL, NULL, (LPBYTE)sValue, &dwSize);
	if (nRet == ERROR_SUCCESS) {
		_stscanf_s(sValue, _T("%hd"), &m_wLanguage);
		if (m_wLanguage == 0) {
			// Will use Office language
			Office2000::LanguageSettings *pSettings = NULL;
			hr = m_pApplication->get_LanguageSettings(&pSettings);
			if (SUCCEEDED(hr)) {
				int office = 0;
				hr = pSettings->get_LanguageID(msoLanguageIDUI, &office);
				if (SUCCEEDED(hr)) {
					m_wLanguage = (WORD)office;
#ifdef _DEBUG
					{
						TCHAR sBuf[256];
						StringCchPrintf(sBuf, 256, _T("Using Office UI Language : %d\n"), m_wLanguage);
						OutputDebugString(sBuf);
					}
#endif
				} else {
					TraceError(hr, _T("Unable to read office language settings"));
				}
				pSettings->Release();
			} else {
				TraceError(hr, _T("Unable to get office language settings"));
			}
		} else {
#ifdef _DEBUG
			{
				TCHAR sBuf[256];
				StringCchPrintf(sBuf, 256, _T("Using Addin specific Language : %d\n"), m_wLanguage);
				OutputDebugString(sBuf);
			}
#endif
		}
	} else {
#ifdef _DEBUG
		{
			TCHAR sBuf[256];
			StringCchPrintf(sBuf, 256, _T("Failed to read language option: %d\n"), nRet);
			OutputDebugString(sBuf);
		}
#endif
	}

	InstallMessageFilter();
	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection(AddInDesignerObjects::ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/ )
{
#ifdef _DEBUG
	OutputDebugString(_T("OnDisconnection\n"));
#endif

	UninstallMessageFilter();

	m_pApplication = NULL;
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate (SAFEARRAY ** /*custom*/ )
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete (SAFEARRAY ** /*custom*/ )
{
	CSetLocale locale(m_wLanguage);

	_CommandBars *pCommandBars = NULL;
	HRESULT hr = m_pApplication->get_CommandBars(&pCommandBars);
	if (FAILED(hr)) {
		TraceError(hr, _T("get_CommandBars"));
		ShowMessage(IDS_UNABLE_TO_CONFIGURE_MENU);
	} else {
		hr = ConfigureMenus(pCommandBars, IDS_EXPORTLABEL, IDS_IMPORTLABEL);
		if (FAILED(hr)) {
			ShowMessage(IDS_UNABLE_TO_CONFIGURE_MENU);
		}
		pCommandBars->Release();
	}
	return hr;
}

STDMETHODIMP CConnect::OnBeginShutdown (SAFEARRAY ** /*custom*/ )
{
#ifdef _DEBUG
	OutputDebugString(_T("Word2000Addin::OnBeginShutdown\n"));
#endif
	_CommandBars *pCommandBars = NULL;
	HRESULT hr = m_pApplication->get_CommandBars(&pCommandBars);
	if (SUCCEEDED(hr)) {
		VARIANT vtCommandBarName;
		vtCommandBarName.vt = VT_BSTR;
		vtCommandBarName.bstrVal = SysAllocString(L"File");

		CommandBar *pCommandBar = NULL;
		hr = pCommandBars->get_Item(vtCommandBarName, &pCommandBar);
		VariantClear(&vtCommandBarName);

		if (SUCCEEDED(hr)) {
			DeleteButton(pCommandBar, IDS_IMPORTLABEL);
			DeleteButton(pCommandBar, IDS_EXPORTLABEL);
			pCommandBar->Release();
		}
		pCommandBars->Release();
	}

	CleanMenus();

	// Removed from XP and 2003 Addins
	MSWord::Template *pTemplate = NULL;
	hr = m_pApplication->get_NormalTemplate(&pTemplate);
	if (FAILED(hr)) {
		TraceError(hr, _T("get_NormalTemplate()"));
	} else {
		hr = pTemplate->Save();
		if (FAILED(hr)) {
			TraceError(hr, _T("pTemplate->Save()"));
		} else {
#ifdef _DEBUG
			OutputDebugString(_T("NormalTemplate saved\n"));
#endif
		}
	}


	// Notify the converter server to shuwdown
	m_converter.Release();

	// TODO : cleanup other variables
	return S_OK;
}


HRESULT CConnect::Clicked(int id)
{
	CSetLocale locale(m_wLanguage);
#ifdef _DEBUG
	OutputDebugString(_T("Clicked\n"));
#endif
	switch (id) {
	case ID_IMPORT:
		return OdtToDocx();
	case ID_EXPORT:
		return DocxToOdt();
	}
	return S_OK;
}



HWND CConnect::FindWord()
{
	HWND hFocus = GetFocus();
	HWND hParent;
	while ((hParent = GetParent(hFocus)) != NULL) {
		hFocus = hParent;
	}
#ifdef _DEBUG
	TCHAR sBuf[512];
	StringCchPrintf(sBuf, 512, _T("Après remontée = 0x%08X\n"), hFocus);
	OutputDebugString(sBuf);
#endif
	return hFocus;
}

int CConnect::ComputeCenter(HWND hWnd)
{
	RECT rect;
	GetWindowRect(hWnd, &rect);
	short centerX = (short)((rect.left + rect.right) / 2);
	short centerY = (short)((rect.top + rect.bottom) / 2);
	return centerX * 65536 + centerY;
}



// Conversion from Odt to Word
HRESULT CConnect::OdtToDocx()
{
	// Variables to cleanup before exit
	_Document *pNewDoc = NULL;
	Fields *pFields = NULL;
	BSTR bstrTempFile = NULL;
	BSTR bstrInputFile = NULL;

	HRESULT hr = S_OK;
	HWND hWordWindow = FindWord();

	TCHAR sFilter[512];
	LoadString(IDS_FILE_FILTER, sFilter, 512);
	// prepare filter
	int len = (int)_tcslen(sFilter);
	for (int i=0;i<len;i++) {
		if (sFilter[i] == '|') {
			sFilter[i] = 0;
		}
	}
	TCHAR sTitle[512];
	LoadString(IDS_IMPORT_TITLE, sTitle, 512);

	TCHAR sFilename[MAX_PATH+1];
	sFilename[0] = 0;

	OPENFILENAME ofn;
	ZeroMemory(&ofn, sizeof(ofn));
	ofn.hwndOwner = hWordWindow;
	ofn.lStructSize = sizeof(ofn);
	ofn.lpstrFilter = sFilter;
	ofn.lpstrFile = sFilename;
	ofn.nMaxFile = MAX_PATH;
	ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_EXPLORER | OFN_ENABLEHOOK;
	ofn.lpfnHook = MyHookProc;
	ofn.lpstrTitle = sTitle;
	if (GetOpenFileName(&ofn)) {
		// Now we must convert the document
		// 1. Ensure that the real converter is instanciated
		HRESULT hr = m_converter.CreateInstance();
		if (FAILED(hr)) {
			ShowMessage(IDS_UNABLE_TO_LAUNCH_CONVERTER);
			goto Cleanup;
		}

		TCHAR sTempFile[MAX_PATH+1];
		if (!GenerateTempName(sTempFile, MAX_PATH, PathFindFileName(sFilename), _T(".docx"))) {
			ShowMessage(IDS_FAILED_TO_OPEN_DOCUMENT);
			goto Cleanup;
		}
#ifdef _DEBUG
		OutputDebugString(_T("Temporary Word12 document : "));
		OutputDebugString(sTempFile);
		OutputDebugString(_T("\n"));
#endif


		bstrTempFile = SysAllocString(sTempFile);
		bstrInputFile = SysAllocString(sFilename);

		hr = m_converter.OdtToDocx(bstrInputFile, bstrTempFile, TRUE, m_wLanguage, ComputeCenter(hWordWindow));
		if (FAILED(hr)) {
			TraceError(hr, _T("Appel : OdfToOox"));
			goto Cleanup;
		}

		// Open the document in Word
		VARIANT vtReadonly;
		vtReadonly.vt = VT_BOOL;
		vtReadonly.boolVal = VARIANT_TRUE;

		OutputDebugString(_T("Opening document\n"));
		VARIANT vtDocument;
		vtDocument.vt = VT_BSTR;
		vtDocument.bstrVal = bstrTempFile;

		hr = m_pDocuments->Open(&vtDocument, &vtMissing, &vtReadonly, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &pNewDoc);
		if (FAILED(hr)) {
			TraceError(hr, _T("Documents->Open()"));
			ShowMessage(IDS_FAILED_TO_OPEN_DOCUMENT);
			goto Cleanup;
		}

		/*clam, dialogika: task 1837107
		hr = pNewDoc->get_Fields(&pFields);
		if (SUCCEEDED(hr)) {
			long nbUpdated = 0;
			pFields->Update(&nbUpdated);
		}*/

		// Finally activate the document
		hr = pNewDoc->Activate();
		if (FAILED(hr)) {
			TraceError(hr, _T("pDoc->Activate()"));
			ShowMessage(IDS_FAILED_TO_OPEN_DOCUMENT);
			goto Cleanup;
		}

	} else {
		DWORD dwErr = CommDlgExtendedError();
		if (dwErr == 0) {
			// Cancel
		} else {
			TraceError(dwErr, _T("GetOpenFilename"));
		}
	}

Cleanup:
	if (bstrTempFile) SysFreeString(bstrTempFile);
	if (bstrInputFile) SysFreeString(bstrInputFile);
	if (pNewDoc) pNewDoc->Release();
	if (pFields) pFields->Release();
	return hr;
}



// Conversion from Word to Odt
HRESULT CConnect::DocxToOdt()
{
	OutputDebugString(_T("DocxToOdt\n"));

	// Variables to cleanup before exit
	_Document *pCurrentDoc = NULL;
	_Document *pNewDocument = NULL;
	BSTR bstrFullname = NULL;
	BSTR bstrTempCopyName = NULL;
	BSTR bstrTempDocxName = NULL;
	BSTR bstrTargetFilename = NULL;

	BOOL bDelete = FALSE;

	HRESULT hr = S_OK;
	HWND hWordWindow = FindWord();

	if (m_pApplication == NULL) {
		hr = E_FAIL;
		goto Cleanup;
	}
	// 0. Ensure we know Word 2007 Format
	if (_Word12SaveFormat == -1) {
		hr = GetWord2007SaveFormat();
		if (FAILED(hr)) {
			ShowMessage(IDS_NOWORD12);
			hr = S_OK;
			goto Cleanup;
		}
	}

	// 1. Ensure that the document is saved
	hr = m_pApplication->get_ActiveDocument(&pCurrentDoc);
	if (FAILED(hr)) {
		hr = E_FAIL;
		goto Cleanup;
	}

	VARIANT_BOOL saved;
	hr = pCurrentDoc->get_Saved(&saved);
	if (FAILED(hr)) {
		ShowMessage(IDS_SAVE_DOC_FIRST);
		hr = E_FAIL;
		goto Cleanup;
	}
	if (saved == VARIANT_FALSE) {
		ShowMessage(IDS_SAVE_DOC_FIRST);
		hr = S_OK;
		goto Cleanup;
	}

	hr = pCurrentDoc->get_FullName(&bstrFullname);
	if (FAILED(hr) || (bstrFullname == NULL) || (_tcschr((LPCWSTR)bstrFullname, L'.') == NULL)) {
		// The document is not saved, ask the user to save it first
		ShowMessage(IDS_SAVE_DOC_FIRST);
		hr = S_OK;
		goto Cleanup;
	}
#ifdef _DEBUG
	OutputDebugString(_T("Original filename : "));OutputDebugString(bstrFullname);OutputDebugString(_T("\n"));
#endif

	// 1. Ask the filename to the user
	TCHAR sFilter[512];
	LoadString(IDS_FILEFILTER, sFilter, 512);
	int len = (int)_tcslen(sFilter);
	for (int i=0;i<len;i++) {
		if (sFilter[i] == '|') {
			sFilter[i] = 0;
		}
	}

	TCHAR sTitle[512];
	LoadString(IDS_EXPORTTITLE, sTitle, 512);

	TCHAR sFilename[MAX_PATH+1];
	LPCTSTR sOriginalName = PathFindFileName(bstrFullname);
	if (sOriginalName != NULL) {
		StringCchCopy(sFilename, MAX_PATH, sOriginalName);
		LPWSTR pExtension = PathFindExtension(sFilename);
		if (pExtension) {
			// Add the extention
			StringCchCopy(pExtension, MAX_PATH - (pExtension - sFilename), _T(".odt"));
		}

	} else {
		sFilename[0] = 0;
	}

	OPENFILENAME ofn;
	ZeroMemory(&ofn, sizeof(ofn));
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hWordWindow;
	ofn.lpstrFilter = sFilter;
	ofn.lpstrFile = sFilename;
	ofn.nMaxFile = MAX_PATH;
	ofn.lpstrTitle = sTitle;
	ofn.Flags = OFN_EXPLORER | OFN_ENABLEHOOK;
	ofn.lpfnHook = MyHookProc;

	if (!GetSaveFileName(&ofn)) {
#ifdef _DEBUG
		DWORD dwErr = CommDlgExtendedError();
		if (dwErr != 0) {
			// Other error
			TraceError(dwErr, _T("GetOpenFilename"));
		}
#endif
		goto Cleanup;
	}

	LPCWSTR pExtension = PathFindExtension(sFilename);
	if (!*pExtension) {
		// Add the extention
		StringCchCat(sFilename, MAX_PATH+1, _T(".odt"));
	}
#ifdef _DEBUG
	OutputDebugString(_T("Target filename : "));OutputDebugString(sFilename);OutputDebugString(_T("\n"));
#endif

	// Now we must convert the document
	// 1. Ensure that the real converter is instanciated
	hr = m_converter.CreateInstance();
	if (FAILED(hr)) {
		ShowMessage(IDS_UNABLE_TO_LAUNCH_CONVERTER);
		goto Cleanup;
	}

	// 1. Ensure the document is saved with Office 2007 Format
	long currentFormat;
	hr = pCurrentDoc->get_SaveFormat(&currentFormat);
	if (FAILED(hr)) {
		ShowMessage(IDS_INTERNAL_ERROR);
		goto Cleanup;
	}

	if (currentFormat != _Word12SaveFormat) {
#ifdef _DEBUG
		OutputDebugString(_T("Need to create a copy in Word12 Format\n"));
#endif

		// => Need create a Word 12 document first
		// 1.1 Copy the current file
		TCHAR sTempCopyName[MAX_PATH+1];
		if (!GenerateTempName(sTempCopyName, MAX_PATH, PathFindFileName((LPCWSTR)bstrFullname), PathFindExtension((LPCWSTR)bstrFullname))) {
#ifdef _DEBUG
			OutputDebugString(_T("Unable to create copy temp name\n"));
#endif
			ShowMessage(IDS_EXPORT_TRY_DOCX);
			goto Cleanup;
		}
#ifdef _DEBUG
		OutputDebugString(_T("Temp Copy file : "));OutputDebugString(sTempCopyName);OutputDebugString(_T("\n"));
#endif

		if (!CopyFile((LPCWSTR)bstrFullname, sTempCopyName, TRUE)) {
#ifdef _DEBUG
			OutputDebugString(_T("Unable to create copy of file\n"));
#endif
			ShowMessage(IDS_EXPORT_TRY_DOCX);
			goto Cleanup;
		}

		// 1.2 Open this document in background
		VARIANT vtFalse;
		vtFalse.vt = VT_BOOL;
		vtFalse.boolVal = VARIANT_FALSE;

		bstrTempCopyName = SysAllocString(sTempCopyName);
		VARIANT vtFileName;
		vtFileName.vt = VT_BSTR;
		vtFileName.bstrVal = bstrTempCopyName;

		hr = m_pDocuments->Open(&vtFileName, &vtMissing, &vtMissing, &vtFalse, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtMissing, &vtFalse, &pNewDocument);
		if (FAILED(hr)) {
			TraceError(hr, _T("pDocuments::Open"));
			ShowMessage(IDS_EXPORT_TRY_DOCX);
			goto Cleanup;
		} else {
#ifdef _DEBUG
			OutputDebugString(_T("Copy document opened\n"));
#endif
		}

		// 1.3 Save this document with .docx format
		TCHAR sTempDocxName[MAX_PATH+1];
		if (!GenerateTempName(sTempDocxName, MAX_PATH, PathFindFileName((LPCWSTR)bstrFullname), _T(".docx"))) {
			ShowMessage(IDS_EXPORT_TRY_DOCX);
			goto Cleanup;
		}
#ifdef _DEBUG
		OutputDebugString(_T("Temp Docx file : "));OutputDebugString(sTempDocxName);OutputDebugString(_T("\n"));
#endif

		bstrTempDocxName = SysAllocString(sTempDocxName);
		vtFileName.vt = VT_BSTR;
		vtFileName.bstrVal = bstrTempDocxName;

		VARIANT vtWord12;
		vtWord12.vt = VT_I4;
		vtWord12.lVal = _Word12SaveFormat;


		hr = pNewDocument->SaveAs(&vtFileName, &vtWord12, &vtMissing, &vtMissing, &vtFalse);
		if (FAILED(hr)) {
			TraceError(hr, _T("pDocuments->SaveAs"));
			ShowMessage(IDS_EXPORT_TRY_DOCX);
			goto Cleanup;
		} else {
#ifdef _DEBUG
			OutputDebugString(_T("Docx file saved\n"));
#endif
		}

		// 1.4 Close and remove the duplicated file
		pNewDocument->Close(&vtFalse);
		DeleteFile(sTempCopyName);

		SysFreeString(bstrFullname);
		bstrFullname = bstrTempDocxName;
		bstrTempDocxName = NULL;
		bDelete = TRUE;
	}

#ifdef _DEBUG
	OutputDebugString(_T("Preparing conversion launch : "));OutputDebugString(bstrFullname);OutputDebugString(_T("\n"));
#endif

	// 2. Launch the conversion
	bstrTargetFilename = SysAllocString(sFilename);
	hr = m_converter.DocxToOdt(bstrFullname, bstrTargetFilename, TRUE, m_wLanguage, ComputeCenter(hWordWindow));

	if (bDelete) {
		DeleteFile(bstrFullname);
	}

Cleanup:
	if (pCurrentDoc) pCurrentDoc->Release();
	if (pNewDocument) pNewDocument->Release();
	if (bstrFullname) SysFreeString(bstrFullname);
	if (bstrTempCopyName) SysFreeString(bstrTempCopyName);
	if (bstrTempDocxName) SysFreeString(bstrTempDocxName);
	if (bstrTargetFilename) SysFreeString(bstrTargetFilename);

	return hr;
}


HRESULT CConnect::DeleteButton(CommandBar *pCommandBar, UINT idResource)
{
#ifdef _DEBUG
	OutputDebugString(_T("OdfOffice2000Addin::DeleteButton\n"));
#endif
	if (pCommandBar == NULL) {
#ifdef _DEBUG
		OutputDebugString(_T("Called after pCommandBar has been cleared\n"));
#endif
		return E_FAIL;
	} else {
#ifdef _DEBUG
		TCHAR sTmp[128];
		StringCchPrintf(sTmp, 128, _T("pCommandBar = 0x%08X\n"), pCommandBar);
		OutputDebugString(sTmp);
#endif
	}
	int old = GetThreadLocale();
	SetThreadLocale(m_wLanguage);

	// Variables to cleanup before exit
	VARIANT vtTag;
	CommandBarControl *pButton = NULL;

	// Register Import label and command
	TCHAR sTag[256];
	LoadString(idResource, sTag, 256);
	vtTag.vt = VT_BSTR;
	vtTag.bstrVal = SysAllocString(sTag);;

#ifdef _DEBUG
	OutputDebugString(_T("calling FindControl()\n"));
	OutputDebugString(vtTag.bstrVal);
	OutputDebugString(_T("\n"));
#endif
	HRESULT hr = pCommandBar->FindControl(vtMissing, vtMissing, vtTag, vtMissing, vtMissing, &pButton);
	if (FAILED(hr) || (pButton == NULL)) {
		TraceError(hr, _T("pCommandBar->FindControl()"));
	} else {
#ifdef _DEBUG
		OutputDebugString(_T("calling pButton->Delete()\n"));
#endif
		hr = pButton->Delete();
		if (FAILED(hr)) {
			TraceError(hr, _T("pButton->Delete()"));
		} else {
#ifdef _DEBUG
			OutputDebugString(_T("pButton->Delete() OK\n"));
#endif
		}
		pButton->Release();
#ifdef _DEBUG
		OutputDebugString(_T("pButton->Release() OK\n"));
#endif
	}


Cleanup:
	VariantClear(&vtTag);
	SetThreadLocale(old);
	return hr;
}



