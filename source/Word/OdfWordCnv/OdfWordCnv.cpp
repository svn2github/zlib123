// OdfWordCnv.cpp : Defines the entry point for the DLL application.
//
#using <mscorlib.dll>
#using <system.dll>
#using <system.drawing.dll>
#using <system.windows.forms.dll>

// ODF converter is loaded dynamically 
//#using "OdfAddinLib.dll"
//#using "OdfConverterLib.dll"
//#using "WordprocessingConverter.dll"
//#using "OdfWordAddin.dll"

#include <windows.h>
#include <objidl.h>

#include "convapi.h"
#include "debug.h"
#include "util.h"
#include "resource.h"

#include "OdfWordCnv.h"

#pragma unmanaged;

char szzReadClasses[] = "OdfConverter\0";
char szzWriteClasses[] = "OdfConverter\0";

char szzReadExts[] = "odt\0";
char szzWriteExts[] = "odt\0";	// only one write extension per class

typedef struct converter {
   HINSTANCE				handle;
   InitConverter32Func		*InitConverter;
   UninitConverterFunc		*FreeConverter;
   RegisterAppFunc          *RegisterApp;
   ForeignToRtf32Func       *ForeignToRtf;
   RtfToForeign32Func       *RtfToForeign;
   IsFormatCorrect32Func    *IsFormatCorrect;
   CchFetchLpszErrorFunc    *FetchError;
   GetReadNamesFunc         *GetRead;
   GetWriteNamesFunc        *GetWrite;
} converter; 

HINSTANCE hInstance;	// hInstance for this DLL
BOOL vfInit = fFalse;	// global init flag - set at first init
LFREGAPP lfRegApp;		// RegisterApp lflags
HANDLE haszOpenFileName;

struct converter cnv;

const char *WORD12_ERROR = "Word12 converter not found. Make sure Office Compatibility Pack is correctly installed.";


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
    return TRUE;
}



/* I N I T  C O N V E R T E R  3  2 */
/*----------------------------------------------------------------------------
	%%Function: InitConverter32

	Required entrypoint.  Preserve hwndCaller or szModule if you care to
	use them later.  Provide a DllMain entrypoint and preserve its
	hInstance argument if you care to use that value later, or to protect
	the DLL from multiple callers if not re-entrant.

	Returns:
	non-zero if all went well.
----------------------------------------------------------------------------*/
LONG PASCAL InitConverter32(HANDLE hwndCaller, char *szModule)
{
    int  ret;
	HKEY hKey;
	char path[MAX_PATH];
	DWORD dwSize;

	memset (&cnv, 0, sizeof (struct converter));
    
	// TODO: Retrieve Path from Registry
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE,
        "SOFTWARE\\Microsoft\\Shared Tools\\Text Converters\\Export\\Word12", 0, KEY_READ, &hKey) != ERROR_SUCCESS)
    {
        MessageBox(NULL, WORD12_ERROR, "Error", MB_ICONERROR | MB_OK | MB_DEFBUTTON1);
		return fFalse;
    }

	if (RegQueryValueEx(hKey, "Path", NULL, NULL, (LPBYTE)&path, &dwSize) != ERROR_SUCCESS)
    {
		MessageBox(NULL, WORD12_ERROR, "Error", MB_ICONERROR | MB_OK | MB_DEFBUTTON1);
		return fFalse;
	}

    cnv.handle = LoadLibrary (path);
    if (cnv.handle == (HINSTANCE)HINSTANCE_ERROR) {
        MessageBox(NULL, WORD12_ERROR, "Error", MB_ICONERROR | MB_OK | MB_DEFBUTTON1);
		return fFalse;
    }

    cnv.InitConverter = (InitConverter32Func*) GetProcAddress (cnv.handle, "InitConverter32");
    cnv.IsFormatCorrect = (IsFormatCorrect32Func*) GetProcAddress (cnv.handle, "IsFormatCorrect32");
    cnv.ForeignToRtf = (ForeignToRtf32Func*) GetProcAddress (cnv.handle, "ForeignToRtf32");
    cnv.RtfToForeign = (RtfToForeign32Func*) GetProcAddress (cnv.handle, "RtfToForeign32");
    cnv.FetchError = (CchFetchLpszErrorFunc*) GetProcAddress (cnv.handle, "CchFetchLpszError");
    cnv.FreeConverter = (UninitConverterFunc*) GetProcAddress (cnv.handle, "UninitConverter");
    cnv.GetRead = (GetReadNamesFunc*) GetProcAddress (cnv.handle, "GetReadNames");
    cnv.GetWrite = (GetWriteNamesFunc*) GetProcAddress (cnv.handle, "GetWriteNames");
    cnv.RegisterApp = (RegisterAppFunc*) GetProcAddress (cnv.handle, "RegisterApp");
    
    if (!cnv.InitConverter) {
        FreeLibrary (cnv.handle);
        MessageBox(NULL, WORD12_ERROR, "Error", MB_ICONERROR | MB_OK | MB_DEFBUTTON1);
		return fFalse;
    }

    ret = cnv.InitConverter ((HANDLE)NULL, szModule);
    if (ret != 1) {
        FreeLibrary (cnv.handle);
        MessageBox(NULL, WORD12_ERROR, "Error", MB_ICONERROR | MB_OK | MB_DEFBUTTON1);
		return ret;
    }

    return fTrue;
}


/* U N I N I T  C O N V E R T E R */
/*----------------------------------------------------------------------------
	%%Function: UninitConverter

	Optional entrypoint.  Uninitialize whatever you need to.
----------------------------------------------------------------------------*/
void PASCAL UninitConverter(void)
{
	if (cnv.FreeConverter) 
		cnv.FreeConverter ();
    FreeLibrary (cnv.handle);
	return;
}

/* I S  F O R M A T  C O R R E C T  3  2 */
/*----------------------------------------------------------------------------
	%%Function: IsFormatCorrect32

	Purpose:
	Return fTrue if ghszFile refers to a file in a format understood by this
	converter.  If this converter deals with multiple formats, return the
	specific format of this file in *ghszClass.  Return an empty
	string in *ghszClass if version is unknown.

	Parameters:
	ghszFile :	global handle to '\0'-terminated filename
	ghszClass :	global handle to return file's version

	Note:
	Currently, no pstgForeign argument here.  This is unfortunate and may
	be fixed.

	Returns:
	0 if format not recognized, 1 if format recognized, fce if error
----------------------------------------------------------------------------*/
FCE PASCAL IsFormatCorrect32(HANDLE ghszFile, HANDLE ghszClass)
{
	/*HANDLE hFile;
	char szFileName[260];
	char rgb[cchFilePrefix];
	long cbr;

	// copy filename locally; file open doesn't want a global handle
	AssertSz(ghszFile != NULL, "NULL filename in IsFormatCorrect");
	if (!PchBltSzHx(szFileName, ghszFile))
		return fFalse;
	// translate filename from the oem character set to the windows set
	if (!lfRegApp.fDontNeedOemConvert)
		OemToChar((LPCSTR)szFileName, (LPSTR)szFileName);

	// open file, read sufficient bytes to identify file type, and close.
	if ((hFile = CreateFile((LPSTR)szFileName, GENERIC_READ,
						FILE_SHARE_READ, (LPSECURITY_ATTRIBUTES)0,
						OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN,
						(HANDLE)NULL)) == INVALID_HANDLE_VALUE)
		{
		return fceOpenInFileErr;
		}
	ReadFile(hFile, rgb, cchFilePrefix, &cbr, NULL);
	CloseHandle(hFile);

	// if the file begins with our szFilePrefix, it's most likely in our
	// file format
	if (strncmp(szFilePrefix, rgb, cchFilePrefix) != 0)
		return fFalse;

	// there is only one read class
	HxBltHxSz(ghszClass, szzReadClasses);*/

	return fTrue;
}

/* F O R E I G N  T O  R T F  3  2 */
/*----------------------------------------------------------------------------
	%%Function: ForeignToRtf32

	Purpose:
	Convert a file to Rtf using the format specified in ghszClass.

	Parameters:
	ghszFile :	global handle to '\0'-terminated filename to be read
	pstgForeign : pointer to IStorage of embedding being converted,
				NULL for non-OLE2 docfile converters.
	ghBuff :	global handle to buffer in which chunks of Rtf are passed
				to WinWord.
	ghszClass:	identifies which file format to translate.  This
				string is the one selected by the user in the 'Confirm
				Conversions' dialog box in Word.
	ghszSubset:	identifies which subset of the file is to be converted.
				Typically used by spreadsheet converters to identify
				subranges of the entire spreadsheet to import.
	lpfnOut:	callback function provided by WinWord to be called whenever
				we have a chunk of Rtf to return.

	Returns:
	fce indicating success or failure and cause
----------------------------------------------------------------------------*/
FCE PASCAL ForeignToRtf32(HANDLE ghszFile, IStorage *pstgForeign, HANDLE ghBuff, HANDLE ghszClass, HANDLE ghszSubset, PFN_RTF lpfnOut)
{
	HGLOBAL hClass;
	char szFileName[260];
	int ret;
    
	AssertSz(((ghszFile == NULL) ^ (pstgForeign == NULL)),
		"Need exactly one of ghszFile and pstgForeign");

	if (pstgForeign)
		return fceInvalidDoc;

	// copy filename locally; file open doesn't want a global handle
	AssertSz(ghszFile != NULL, "NULL filename in ForeignToRtf");
	if (!PchBltSzHx(szFileName, ghszFile))
		return fceNoMemory;

	if (!cnv.ForeignToRtf) {
        return fceOpenConvErr;
    }
    
	ret = Convert(ghszFile, false);
	if (ret < 0) {
        return ret;
    }

	hClass = StringToHGLOBAL("Word12");
	ret = cnv.ForeignToRtf(haszOpenFileName, pstgForeign, ghBuff, hClass, ghszSubset, lpfnOut);
	GlobalFree (hClass);

	return ret;
}

/* R T F  T O  F O R E I G N  3  2 */
/*----------------------------------------------------------------------------
	%%Function: RtfToForeign32

	Purpose:
	Convert a stream of Rtf from Word to a file in the format specified in
	ghszClass.

	Parameters:
	ghszFile :	global handle to '\0'-terminated filename to be written
	ghBuff :	global handle to buffer in which chunks of Rtf are passed
				to us from WinWord.
	ghszClass:  identifies which file format to translate to.  This
				string is the one selected by the user in the 'Save File
				as Type' drop down in Word's Save dialog.
	lpfnIn:		callback function provided by WinWord to be called whenever
				we are ready for more Rtf.

	Returns:
	fce indicating success or failure and cause
----------------------------------------------------------------------------*/
FCE PASCAL RtfToForeign32(HANDLE ghszFile, IStorage *pstgForeign, HANDLE ghBuff, HANDLE ghszClass, PFN_RTF lpfnIn)
{
	HGLOBAL hClass;
	char szFileName[260];
	int ret;
    
	AssertSz(((ghszFile == NULL) ^ (pstgForeign == NULL)),
		"Need exactly one of ghszFile and pstgForeign");

	if (pstgForeign)
		return fceInvalidDoc;

	// copy filename locally; file open doesn't want a global handle
	AssertSz(ghszFile != NULL, "NULL filename in RtfToForeign");
	if (!PchBltSzHx(szFileName, ghszFile))
		return fceNoMemory;

	if (!cnv.RtfToForeign) {
        return fceOpenConvErr;
    }

	hClass = StringToHGLOBAL("Word12");
    ret = cnv.RtfToForeign (ghszFile, pstgForeign, ghBuff, hClass, lpfnIn);
	GlobalFree (hClass);

    if (ret < 0) {
        return ret;
    }

	return Convert(ghszFile, true);
}

/* G E T  R  W  N A M E S */
/*----------------------------------------------------------------------------
	%%Function: GetRWNames

	Do work for GetReadNames and GetWriteNames
----------------------------------------------------------------------------*/
VOID GetRWNames(char** haszClass, char** haszDescrip, char** haszExt, BOOL fRead)
{
	char *szzClasses;
	char *szzExts;
	char szzNames[cbResMax];
	
	AssertSz(haszClass && GlobalSize(haszDescrip) >= 2,
		"Null/small handle for class names. About to crash");
	AssertSz(haszDescrip && GlobalSize(haszDescrip) >= 2,
		"Null/small handle for descriptions. About to crash.");
	AssertSz(haszExt && GlobalSize(haszExt) >= 2,
		"Null/small handle for extensions. About to crash.");

	if (fRead)
		{
		szzClasses = szzReadClasses;
		szzExts = szzReadExts;
		AssertDo(LoadString(hInstance, rcidReadNames, szzNames, cbResMax));
		SzzFromSoz(szzNames, szzNames);
		}
	else
		{
		szzClasses = szzWriteClasses;
		szzExts = szzWriteExts;
		AssertDo(LoadString(hInstance, rcidWriteNames, szzNames, cbResMax));
		SzzFromSoz(szzNames, szzNames);
		}

	if (!HxBltHxSzz(haszClass, szzClasses)
	 || !HxBltHxSzz(haszDescrip, szzNames)
	 || !HxBltHxSzz(haszExt, szzExts))
		{
		// if any copy fails, return none of the values
		*haszClass = '\0';
		*haszDescrip = 0;
		*haszExt = 0;
		}
}


/* G E T  R E A D  N A M E S */
/*----------------------------------------------------------------------------
	%%Function: GetReadNames

	Required entrypoint.  Returns import class names, description strings
	and file extensions for this converter.
----------------------------------------------------------------------------*/
void PASCAL GetReadNames(HANDLE haszClass, HANDLE haszDescrip, HANDLE haszExt)
{
	GetRWNames((char**)haszClass, (char**)haszDescrip, (char**)haszExt, fTrue/*fRead*/);
}


/* G E T  W R I T E  N A M E S */
/*----------------------------------------------------------------------------
	%%Function: GetWriteNames

	Required entrypoint.  Returns export class names, description strings
	and file extensions for this converter.
----------------------------------------------------------------------------*/
void PASCAL GetWriteNames(HANDLE haszClass, HANDLE haszDescrip, HANDLE haszExt)
{
	GetRWNames((char**)haszClass, (char**)haszDescrip, (char**)haszExt, fFalse/*fRead*/);
}

/* R E G I S T E R  A P P */
/*----------------------------------------------------------------------------
	%%Function: RegisterApp

	Learn about calling app, and teach it about us.

	Returns:
	handle to a GlobalAlloc'd RegAppRet structure.  Caller must use
	GlobalLock() to access it, and MUST free it if the return value
	is non-NULL.
----------------------------------------------------------------------------*/
HGLOBAL PASCAL RegisterApp(DWORD lFlags, VOID FAR *lpFuture)
{
	HGLOBAL hRegAppRet;
	REGAPPRET FAR *lpRegAppRet;
	
	lfRegApp.lfRegApp = lFlags;

	// Allocate return structure, fill in, return handle
	if ((hRegAppRet = GlobalAlloc(GMEM_MOVEABLE, sizeof(REGAPPRET))) == NULL)
		return (HGLOBAL)0;
		
	lpRegAppRet = (REGAPPRET FAR *)GlobalLock(hRegAppRet);
	Assert(lpRegAppRet != NULL);

	lpRegAppRet->cbStruct = sizeof(REGAPPRET);

	// docfile record
	lpRegAppRet->cbSizefDocfile = sizeof(CHAR)+sizeof(CHAR)+sizeof(SHORT);
	lpRegAppRet->opcodefDocfile = RegAppOpcodeDocfile;
	lpRegAppRet->grfType = 0;		// clear all flags ...
#ifndef WW6EXP
	lpRegAppRet->fNonDocfile = 1;	// ... then set nondocfile
#endif
	
	// Rtf version record
	lpRegAppRet->cbSizeVer = sizeof(CHAR)+sizeof(CHAR)+sizeof(SHORT)+sizeof(SHORT);
	lpRegAppRet->opcodeVer = RegAppOpcodeVer;
	lpRegAppRet->verMajor = 7;		// Word 7.0 Rtf compliant converter
	lpRegAppRet->verMinor = 0;
		
	// character set record
	lpRegAppRet->cbSizeCharset = 3 * sizeof(BYTE);
	lpRegAppRet->opcodeCharset = RegAppOpcodeCharset;
	lpRegAppRet->charset = ANSI_CHARSET;
	
	GlobalUnlock(hRegAppRet);

	return hRegAppRet;
}

/* F  R E G I S T E R  C O N V E R T E R */
/*----------------------------------------------------------------------------
	%%Function: FRegisterConverter

	Register converter in appropriate 'Text Converters' section.
----------------------------------------------------------------------------*/
LONG PASCAL FRegisterConverter(HANDLE hkeyRoot)
{
	char szzNames[cbResMax];
	char szModulePath[cbSzPathMax];

 	if (!GetModuleFileName(hInstance, (LPTSTR)szModulePath, cbSzPathMax - 1))
		{
		Assert(fFalse);
		return fFalse;
		}

	if (*szzReadClasses)
		{
		AssertDo(LoadString(hInstance, rcidReadNames, szzNames, cbResMax));
		SzzFromSoz(szzNames, szzNames);
		if (!FRegisterClasses(hkeyRoot, "Text Converters\\Import",
			szzReadClasses, szzReadExts, szzNames, szModulePath))
			return fFalse;
		}

	if (*szzWriteClasses)
		{
		AssertDo(LoadString(hInstance, rcidWriteNames, szzNames, cbResMax));
		SzzFromSoz(szzNames, szzNames);
		if (!FRegisterClasses(hkeyRoot, "Text Converters\\Export",
			szzWriteClasses, szzWriteExts, szzNames, szModulePath))
			return fFalse;
		}

	return fTrue;
}

/* F  R E G I S T E R  C L A S S E S */
/*----------------------------------------------------------------------------
	%%Function: FRegisterClasses

	Do the work of registering a set of class extensions/name/path values.
----------------------------------------------------------------------------*/
static BOOL FRegisterClasses(HANDLE hkeyRoot, char *szSection,
			char *szzClasses, char *szzExts, char *szzNames, char *szModulePath)
{
	HKEY hkeySection;
	HKEY hkeyCurrent;
	DWORD disp;

	if (RegCreateKeyEx((HKEY)hkeyRoot, szSection, 0L, "", 
						REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, NULL,
						&hkeySection, &disp) != ERROR_SUCCESS)
		return fFalse;

	// for each class in szzClasses, add all relevant values
	while (*szzClasses)
		{
		if (RegCreateKeyEx(hkeySection, szzClasses, 0L, "", 
						   REG_OPTION_NON_VOLATILE, KEY_READ | KEY_WRITE, NULL,
						   &hkeyCurrent, &disp) != ERROR_SUCCESS)
			{
			RegCloseKey(hkeySection);
			return fFalse;
			}

		RegSetValueEx(hkeyCurrent, "Extensions", 0, REG_SZ,
								(BYTE *)szzExts, strlen(szzExts));
		RegSetValueEx(hkeyCurrent, "Name", 0, REG_SZ,
								(BYTE *)szzNames, strlen(szzNames));
		RegSetValueEx(hkeyCurrent, "Path", 0, REG_SZ,
								(BYTE *)szModulePath, strlen(szModulePath));
		
		// now, increment the szz's
		szzClasses += (strlen(szzClasses) + 1);
		szzExts += (strlen(szzExts) + 1);
		szzNames += (strlen(szzNames) + 1);
		
		AssertSz(!szzClasses[0] || szzExts[0], "Not enough extensions");
		AssertSz(!szzClasses[0] || szzNames[0], "Not enough names");
		}

	RegCloseKey(hkeyCurrent);
	RegCloseKey(hkeySection);

	return fTrue;
}

#pragma managed

#undef GetCurrentDirectory // don't expand the macro defined in winbase.h, use .NET function instead

FCE PASCAL Convert(HANDLE ghszFile, bool bExport)
{
	// Using a nested call to Convert2, thus we are able to catch 
	// Exceptions if .NET assemblies cannot be retrieved. 
	// Any uncatched .NET exception will crash winword.exe.
	FCE ret = fceOpenConvErr;
	try
	{
		ret = Convert2(ghszFile, bExport);
	}
	catch (System::Exception ^ex)
	{
		System::Windows::Forms::MessageBox::Show(ex->ToString());
	}
	return ret;
}

FCE PASCAL Convert2(HANDLE ghszFile, bool bExport)
{
	char szFileName[260];
	
	if (!PchBltSzHx(szFileName, ghszFile))
		return fceNoMemory;

	// translate filename from the oem character set to the windows set
	if (!lfRegApp.fDontNeedOemConvert)
		OemToChar((LPCSTR)szFileName, (LPSTR)szFileName);
	
	try
	{
		// call methods dynamically
		// Reason: We are inside the winword.exe process, meaning that the ODF converter assemblies 
		// would have to be intalled to the GAC or to the winword.exe folder or be configured in machine.config
		// in order to be found by the .NET runtime. 
		System::Reflection::Assembly ^thisAssembly = System::Reflection::Assembly::GetExecutingAssembly();
		System::IO::FileInfo ^fi = gcnew System::IO::FileInfo(thisAssembly->Location);

		System::String ^path = System::IO::Path::Combine(fi->DirectoryName, "WordprocessingConverter.dll");
		System::Reflection::Assembly ^asmConverter = System::Reflection::Assembly::LoadFrom(path);
		System::Type ^convType = asmConverter->GetType("OdfConverter.Wordprocessing.Converter");
		System::Object ^objConverter = System::Activator::CreateInstance(convType);
		
		path = System::IO::Path::Combine(fi->DirectoryName, "OdfWordAddin.dll");
		System::Reflection::Assembly ^asmAddin = System::Reflection::Assembly::LoadFrom(path);
		System::Type ^addinType = asmAddin->GetType("OdfConverter.Wordprocessing.OdfWordAddin.Connect");
		System::Object ^objAddin = System::Activator::CreateInstance(addinType);
		
		path = System::IO::Path::Combine(fi->DirectoryName, "OdfAddinLib.dll");
		System::Reflection::Assembly ^asmAddinLib = System::Reflection::Assembly::LoadFrom(path);
		System::Type ^addinLibType = asmAddinLib->GetType("CleverAge.OdfConverter.OdfConverterLib.OdfAddinLib");
		System::Object ^objAddinLib = System::Activator::CreateInstance(addinLibType, objAddin, objConverter);
		
		//OdfConverter::Wordprocessing::Converter ^conv = gcnew OdfConverter::Wordprocessing::Converter();
		//OdfConverter::Wordprocessing::OdfWordAddin::Connect ^connect = gcnew OdfConverter::Wordprocessing::OdfWordAddin::Connect();
		//OdfAddinLib ^addinLib = gcnew OdfAddinLib(connect, conv);

		if (bExport)
		{
			System::String ^tempFile = System::IO::Path::GetTempFileName();
			System::String ^outputFile = gcnew System::String(szFileName);

			System::Reflection::MethodInfo ^mi = addinLibType->GetMethod("OoxToOdf");
			array<System::Object^>^ args = { outputFile, tempFile, true };
			mi->Invoke(objAddinLib, args);
			//addinLib->OoxToOdf(outputFile, tempFile, true);

			System::IO::File::Delete(outputFile);
			System::IO::File::Move(tempFile, outputFile);
		}
		else
		{
			System::String ^tempFile = System::IO::Path::GetTempFileName();
			System::String ^inputFile = gcnew System::String(szFileName);
			
			System::Reflection::MethodInfo ^mi = addinLibType->GetMethod("OdfToOox");
			array<System::Object^>^ args = { inputFile, tempFile, true };
			mi->Invoke(objAddinLib, args);
			//addinLib->OdfToOox(inputFile, tempFile, true);
			
			haszOpenFileName = static_cast<HANDLE>(const_cast<void*>(static_cast<const
				void*>(System::Runtime::InteropServices::Marshal::StringToHGlobalAnsi(tempFile))));
		}

	}
	catch (System::Exception ^ex)
	{
		System::Windows::Forms::MessageBox::Show(ex->ToString());
	}

}
