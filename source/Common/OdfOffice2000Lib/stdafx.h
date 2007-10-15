// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#ifndef _WIN32_WINNT		// Allow use of features specific to Windows XP or later.                   
#define _WIN32_WINNT 0x0501	// Change this to the appropriate value to target other versions of Windows.
#endif						

#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers



#include <tchar.h>
#include <atlbase.h>
#include <atlcom.h>

#import "lib\MSO9.DLL" rename("RGB","OfficeRGB"),rename_namespace("Office2000"),rename("DocumentProperties","O2000DocumentProperties"), raw_interfaces_only
using namespace Office2000;
#import "lib\VBE6EXT.olb" rename_namespace("VBE6"),raw_interfaces_only
using namespace VBE6;
//#import "lib\MSWORD9.OLB" rename("ExitWindows","O2000ExitWindows"),rename("Fonts","O2000Fonts"),rename("FindText","O2000FindText"),named_guids,rename_namespace("MSWord"),raw_interfaces_only
//using namespace MSWord;

#include <strsafe.h>
#include <shlwapi.h>


