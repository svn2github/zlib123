#include "stdafx.h"

extern "C" {

	// This function displays an error message
	void DisplayError(DWORD dwErr, LPCTSTR sTitle) {
		TCHAR sBuf[512];
		DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
		if (!FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, dwErr, 0, sBuf, dwBufSize, NULL)) {
			wsprintf(sBuf, _T("err 0x%08X"), dwErr);
		}
		MessageBox(NULL, sBuf, sTitle, MB_OK);
	}

	static LPCTSTR MyEventName = _T("{78FF3DDE-04DF-44eb-9FFB-EA6468418DCD}");
	static LPCTSTR ConcurrentInstallProperty = _T("CONCURRENTINSTALL");

	UINT __stdcall ForceUniqueInstall(MSIHANDLE hInstall)
	{
		// Clear last error
		SetLastError(0);
		HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, MyEventName);
		if ((hEvent != NULL) && (GetLastError() == ERROR_ALREADY_EXISTS)) {
			// Close the handle as we don't need it anymore
			CloseHandle(hEvent);

			// And set the property
			MsiSetProperty(hInstall, ConcurrentInstallProperty, _T("True"));
		}
		return ERROR_SUCCESS;
	}

	typedef struct _ProductDetection {
		LPCTSTR ProductCode;
		LPCTSTR ProductProperty;
	} ProductDetection;

	static ProductDetection ProductsToDetect[] = {
		{	_T("{85B35D52-430E-4890-8214-F65470E0B568}"),	_T("ODFWORD2003")	},
		{	_T("{56DB0072-4AC3-47E6-8C06-B8BB58638B36}"),	_T("ODFWORD2007")	},
		{	_T("{57205490-7627-47DE-B91A-8FD0E3C80EB3}"),	_T("ODFWORDXP")		}
	};

	static LPCTSTR WordVersionProperty = L"WORDVERSION";

	// This function will check for the presence of old versions of the converter
	UINT __stdcall DetectPreviousConverters(MSIHANDLE hInstall) {
		int productsCount = sizeof(ProductsToDetect) / sizeof(ProductsToDetect[0]);
		for (int i=0;i<productsCount;i++) {
			TCHAR sBuf[512];
			DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
			int nRet = MsiGetProductInfo(ProductsToDetect[i].ProductCode, INSTALLPROPERTY_PRODUCTNAME, sBuf, &dwBufSize);
			switch (nRet) {
			case ERROR_SUCCESS:
				// Product is installed
				OutputDebugString(_T("Product detected : "));
				OutputDebugString(ProductsToDetect[i].ProductProperty);
				OutputDebugString(_T("\n"));
				nRet = MsiSetProperty(hInstall, ProductsToDetect[i].ProductProperty, _T("True"));
				if (nRet != ERROR_SUCCESS) {
					DisplayError(nRet, _T("MsiSetProperty"));
				}
				break;
			case ERROR_UNKNOWN_PRODUCT:
				// Product is not installed
				OutputDebugString(_T("Product NOT detected : "));
				OutputDebugString(ProductsToDetect[i].ProductProperty);
				OutputDebugString(_T("\n"));
				break;
			default:
				DisplayError(nRet, _T("MsiGetProductInfo"));
				break;
			}
		}
		return ERROR_SUCCESS;
	}


	UINT __stdcall GetWordVersion(MSIHANDLE hInstall) {
		HKEY hKey;
		UINT nRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, _T("Word.Application\\Curver"), 0, KEY_READ, &hKey);
		if (nRet == ERROR_SUCCESS) {
			TCHAR sBuf[512];
			DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
			DWORD dwType;
			nRet = RegQueryValueEx(hKey, NULL, NULL, &dwType, (LPBYTE)sBuf, &dwBufSize);
			RegCloseKey(hKey);

			if (nRet == ERROR_SUCCESS) {
				if (!_tcsnicmp(_T("Word.Application."), sBuf, 17)) {
					// Fine
					OutputDebugString(_T("Word version detected : "));
					OutputDebugString(sBuf + 17);
					OutputDebugString(_T("\n"));
					nRet = MsiSetProperty(hInstall, WordVersionProperty, sBuf + 17);
					if (nRet != ERROR_SUCCESS) {
						DisplayError(nRet, _T("MsiSetProperty"));
					}
					return ERROR_SUCCESS;
				}
			}
		}
		// In any case except full detection succeeded
		OutputDebugString(_T("No word version detected"));
		MsiSetProperty(hInstall, WordVersionProperty, L"None");
		return ERROR_SUCCESS;
	}

	static LPCTSTR PropertiesTpDump[] = {
		_T("GUDULEINSTALLED"),
		_T("ODFWORD2003"),
		_T("ODFWORD2007"),
		_T("ODFWORDXP"),
	};

	UINT __stdcall DumpProperties(MSIHANDLE hInstall) {
		int productsCount = sizeof(PropertiesTpDump) / sizeof(PropertiesTpDump[0]);
		for (int i=0;i<productsCount;i++) {
			TCHAR sBuf[512];
			DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
			UINT nRet = MsiGetProperty(hInstall, PropertiesTpDump[i], sBuf, &dwBufSize);
			if (nRet == ERROR_SUCCESS) {
				OutputDebugString(PropertiesTpDump[i]);
				OutputDebugString(_T(" : "));
				OutputDebugString(sBuf);
				OutputDebugString(_T("\n"));
			} else {
				DisplayError(nRet, _T("MsiGetProperty"));
			}
		}
		return ERROR_SUCCESS;
	}

	UINT __stdcall LaunchReadme(MSIHANDLE hInstall) {
		TCHAR sBuf[512];
		DWORD dwBufSize = sizeof(sBuf) / sizeof(sBuf[0]);
		UINT nRet = MsiGetProperty(hInstall, _T("TARGETDIR"), sBuf, &dwBufSize);
		if (nRet == ERROR_SUCCESS) {
			ShellExecute(NULL, _T("Open"), _T("Readme.htm"), NULL, sBuf, SW_SHOW);
		} else {
			DisplayError(nRet, _T("MsiGetProperty"));
		}
		return ERROR_SUCCESS;
	}

	UINT __stdcall NgenAssemblies(MSIHANDLE hInstall) {
		TCHAR sTargetDir[MAX_PATH];
		DWORD dwBufSize = sizeof(sTargetDir) / sizeof(sTargetDir[0]);
		UINT nRet = MsiGetProperty(hInstall, _T("TARGETDIR"), sTargetDir, &dwBufSize);
		if (nRet == ERROR_SUCCESS) {
			WIN32_FIND_DATA wfd;
			HANDLE hFind;
			
			TCHAR sNgen[MAX_PATH] = _T("");
			TCHAR sAssemblyFilePath[MAX_PATH] = _T("");
			
			wcscpy_s(sAssemblyFilePath, sizeof(sAssemblyFilePath) / sizeof(TCHAR), sTargetDir);
			wcscat_s(sAssemblyFilePath, sizeof(sAssemblyFilePath) / sizeof(TCHAR), _T("*.dll"));

			SHGetSpecialFolderPath(0, sNgen, CSIDL_WINDOWS, false);
			wcscat_s(sNgen, sizeof(sNgen) / sizeof(TCHAR), _T("\\Microsoft.NET\\Framework\\v2.0.50727\\ngen.exe")); 

			if ((hFind = FindFirstFile(sAssemblyFilePath, &wfd )) != INVALID_HANDLE_VALUE)
			{
				do
				{
					TCHAR sCmdLine[2*MAX_PATH + 200];
					wcscpy_s(sCmdLine, sizeof(sCmdLine) / sizeof(TCHAR), _T("install "));
					wcscat_s(sCmdLine, sizeof(sCmdLine) / sizeof(TCHAR), wfd.cFileName);
					//wcscat_s(sCmdLine, sizeof(sCmdLine) / sizeof(TCHAR), _T(" /queue"));

					ShellExecute(NULL, _T("Open"), sNgen, sCmdLine, sTargetDir, SW_HIDE);
				} while (FindNextFile(hFind, &wfd));
				FindClose (hFind);
			}

		} else {
			DisplayError(nRet, _T("MsiGetProperty"));
		}
		return ERROR_SUCCESS;
	}
} // extern "C"