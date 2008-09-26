/*----------------------------------------------------------------------------
	%%File: CONVAPI.H
	%%Unit: CORE
	%%Contact: smueller@microsoft.com

	This header is distributed as part of the 32-bit Conversions SDK.

	Changes to this header file should be sent to smueller@microsoft.com

	Revision History: (Current=1.02)

	3/14/97 tidy up and remove references to derived types
	3/16/97 MacPPC callback prototype fix
	4/21/97 explicit packing, MacPPC prototype fixes, C++ friendliness
----------------------------------------------------------------------------*/

#ifndef CONVAPI_H
#define CONVAPI_H

#ifdef __cplusplus
extern "C" {
#endif

#include "converr.h"

// Rtf to foreign Percent ConVersion compleTe
#pragma pack(2)
typedef struct _PCVT
	{
	short cbpcvt;				// size of this structure
	short wVersion;				// version # for determining size of struct
	short wPctWord;				// current %-complete according to Word
	short wPctConvtr;			// current %-complete according to the converter
	} PCVT;
#pragma pack()


// Rtf to Foreign User OPTions
#pragma pack(2)
typedef struct _RFUOPT
	{
	union
		{
		short rgf;
		struct
			{
			short fPicture : 1;		// converter wants pictures (binary implied)
			short fLayoutInfo : 1;	// converter wants (slow) layout info
			short fPctComplete : 1;	// converter will provide percent complete
			short : 13;
			};
		};
	} RFUOPT;
#pragma pack()

// Rtf to foreign user options, again
#define fOptPicture				0x0001
#define fOptLayoutInfo			0x0002
#define fOptPctComplete			0x0004

// RegisterApp input flags
#pragma pack(2)
typedef struct _LFREGAPP
	{
	union
		{
		struct
			{
			unsigned long fAcceptPctComp : 1;
			unsigned long fNoBinary : 1;
			unsigned long fPreview : 1;
			unsigned long fDontNeedOemConvert : 1;
			unsigned long fIndexing : 1;
			unsigned long unused : 27;
			};
		unsigned long lfRegApp;
		};
	} LFREGAPP;
#pragma pack()

// RegisterApp input flags, again
#define fRegAppPctComp			0x00000001	// app is prepared to accept percent complete numbers in export
#define fRegAppNoBinary			0x00000002	// app is not prepared to accept binary-encoded picture and other data
#define fRegAppPreview			0x00000004	// converter should display no message boxes or dialogs
#define fRegAppSupportNonOem	0x00000008	// app is prepared to provide non-OEM character set filenames
#define fRegAppIndexing			0x00000010	// converter can omit non-content Rtf

// REGister APP structure (received from client)
#pragma pack(1)
typedef struct _REGAPP {
	short cbStruct;			// size of this structure
	char rgbOpcodeData[];
	} REGAPP;
#pragma pack()

// REGAPP byte opcodes
#define RegAppOpcodeFilename	0x80	// true final filename of exported file
#define RegAppOpcodeInterimPath	0x81	// path being exported to is not the final location

// REGister APP RETurn structure (passed to client)
#pragma pack(1)
typedef struct _REGAPPRET
	{
	short cbStruct;			// size of this structure

	// Does this converter understand docfiles and/or non-docfiles?
	char cbSizefDocfile;
	char opcodefDocfile;
	union
		{
		struct
			{
			short fDocfile : 1;
			short fNonDocfile : 1;
			short : 14;
			};
		short grfType;
		};

	// Version of Word for which converter's Rtf is compliant
	char cbSizeVer;		// == sizeof(char)+sizeof(char)+sizeof(short)+sizeof(short)
	char opcodeVer;
	short verMajor;		// Major version of Word for which Rtf is compliant
	short verMinor;		// Minor version of Word for which Rtf is compliant

#if defined(WIN16) || defined(WIN32)
	// What character set do we want all filenames to be in.
	char cbSizeCharset;
	char opcodeCharset;
	char charset;
#endif

	char opcodesOptional[];		// optional additional stuff
	} REGAPPRET;
#pragma pack()

// REGAPPRET byte opcodes
#define RegAppOpcodeVer				0x01	// Rtf version converter is prepared to accept
#define RegAppOpcodeDocfile			0x02	// converter's ability to handle regular and docfiles
#define RegAppOpcodeCharset 		0x03	// converter is prepared to handle filenames in this character set
#define RegAppOpcodeReloadOnSave	0x04	// app should reload document after exporting
#define RegAppOpcodePicPlacehold	0x05	// app should send placeholder pictures (with size info) for includepicture \d fields
#define RegAppOpcodeFavourUnicode	0x06	// app should output Unicode RTF whenever possible, especially for DBCS; \uc0 is good
#define RegAppOpcodeNoClassifyChars	0x07	// app should not break text runs by character set clasification



// Principal converter entrypoints

// callback type
typedef long (PASCAL *PFN_RTF)();

typedef short FCE;

long pascal InitConverter32 (HANDLE, char *);
void pascal UninitConverter (void);
void pascal GetReadNames (HANDLE, HANDLE, HANDLE);
void pascal GetWriteNames (HANDLE, HANDLE, HANDLE);
HGLOBAL pascal RegisterApp (unsigned long, void *);
FCE pascal IsFormatCorrect32 (HANDLE, HANDLE);
FCE pascal ForeignToRtf32 (HANDLE, void *, HANDLE, HANDLE, HANDLE, PFN_RTF);
FCE pascal RtfToForeign32 (HANDLE, void *, HANDLE, HANDLE, PFN_RTF);
long pascal CchFetchLpszError (long, char *, long);
long pascal FRegisterConverter (HKEY);

typedef long pascal InitConverter32Func (HANDLE, char *);
typedef void pascal UninitConverterFunc (void);
typedef void pascal GetReadNamesFunc (HANDLE, HANDLE, HANDLE);
typedef void pascal GetWriteNamesFunc (HANDLE, HANDLE, HANDLE);
typedef HGLOBAL pascal RegisterAppFunc (unsigned long, void *);
typedef FCE pascal IsFormatCorrect32Func (HANDLE, HANDLE);
typedef FCE pascal ForeignToRtf32Func (HANDLE, void *, HANDLE, HANDLE, HANDLE, PFN_RTF);
typedef FCE pascal RtfToForeign32Func (HANDLE, void *, HANDLE, HANDLE, PFN_RTF);
typedef long pascal CchFetchLpszErrorFunc (long, char *, long);
typedef long pascal FRegisterConverterFunc (HKEY);


#ifdef __cplusplus
}
#endif

#endif // CONVAPI_H
