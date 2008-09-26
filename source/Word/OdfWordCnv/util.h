/*----------------------------------------------------------------------------
	%%File: UTIL.H

	Utility functions and types and such, for all platforms.
----------------------------------------------------------------------------*/

#ifndef UTIL_H
#define UTIL_H

// common constants
#define fFalse  0
#define fTrue   1
#define cbResMax 256
#define cbSzPathMax 256

// operating system's idea of a handle and platform-independent wrappers
// for system heap functions
typedef HGLOBAL HX;
#define HxAllocLcb(lcb)			GlobalAlloc(GMEM_MOVEABLE, (lcb))
#define FreeHx(hx)				GlobalFree(hx)
#define LpvLockHx(hx)			GlobalLock(hx)
#define UnlockHx(hx)			GlobalUnlock(hx)
#define FReallocHx(hx, lcb)		GlobalReAlloc((hx), (lcb), GMEM_MOVEABLE)

// non-standard utility functions
char *PchBltSzHx(char *sz, HX hxsz);
HX HxBltHxSz(HX hxsz, char *sz);
int CbSzz(char *szz);
HX HxBltHxSzz(HX hxszz, char *szz);
char FAR *SzzFromSoz(char FAR *soz, char FAR *szz);

HGLOBAL StringToHGLOBAL (const char* str);


#endif // UTIL_H

