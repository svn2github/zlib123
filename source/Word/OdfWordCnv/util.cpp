/*----------------------------------------------------------------------------
	%%File: UTIL.C

	Utility functions, for all platforms.
----------------------------------------------------------------------------*/

#include <windows.h>

#include "util.h"


/* P C H  B L T  S Z  H X */
/*----------------------------------------------------------------------------
	%%Function: PchBltSzHx

	Purpose:
	copy a string from a global handle to a near buffer

	Parameters:
	sz :		near pointer to destination buffer
	hxsz :		global handle to source '\0'-terminated string

	Returns:
	sz		iff successful
	NULL	iff not successful
----------------------------------------------------------------------------*/
char *PchBltSzHx(char *sz, HX hxsz)
{
	char FAR *lpsz;
	char *szT = sz;

	if ((lpsz = (char FAR *)LpvLockHx(hxsz)) == NULL)
		return NULL;

	while ((*szT++ = *lpsz++) != 0)
		;

	UnlockHx(hxsz);
	return sz;
}


#ifdef MACPPC
// on Windows platforms, this routine is preferable because it doesn't
// require the C runtimes.  Here, we humour existing Windows-centric code.
#define lstrlen strlen
#endif

/* H X  B L T  H X  S Z */
/*----------------------------------------------------------------------------
	%%Function: HxBltHxSz

	Purpose:
	copy a near string to a global memory block, resizing as necessary.

	Parameters:
	hxsz :		global handle to destination buffer
	sz :		near pointer to source '\0'-terminated string

	Returns:
	hxsz	iff successful
	NULL	iff handle operation failed
----------------------------------------------------------------------------*/
HX HxBltHxSz(HX hxsz, char *sz)
{
	char FAR *lpsz;

	if (!FReallocHx(hxsz, lstrlen(sz) + 1))
		return NULL;

	if ((lpsz = (char FAR *)LpvLockHx(hxsz)) == NULL)
		return NULL;

	while ((*lpsz++ = *sz++) != 0)
		;

	UnlockHx(hxsz);
	return hxsz;
}


/* C B  S Z Z */
/*----------------------------------------------------------------------------
	%%Function: CbSzz

	Purpose:
	Returns the number of bytes in szz, a '\0'-delimited sequence of strings
	ending in '\0\0'

	Parameters:
	szz : '\0'-delimited sequence of strings.

	Returns:
	number of bytes in szz, including all '\0's
----------------------------------------------------------------------------*/
int CbSzz(char *szz)
{
	char *pch;

	pch = szz;
	while (*pch || *(pch + 1))
		pch++;

	return (pch - szz + 2);
}


/* H X  B L T  H X  S Z Z */
/*----------------------------------------------------------------------------
	%%Function: HxBltHxSzz

	Purpose:
	copy a near list of strings to a global memory block, resizing as
	necessary.

	Parameters:
	hxszz :		global handle to destination buffer
	szz :		near pointer to source '\0\0'-terminated string

	Returns:
	hx		iff successful
	NULL	iff not successful
----------------------------------------------------------------------------*/
HX HxBltHxSzz(HX hxszz, char *szz)
{
	char FAR *lpch;
	int cbszz = CbSzz(szz);

	if (!FReallocHx(hxszz, cbszz))
		return NULL;

	if ((lpch = (char FAR *)LpvLockHx(hxszz)) == NULL)
		return NULL;

	// copy szz
	while (cbszz--)
		*lpch++ = *szz++;

	UnlockHx(hxszz);
	return hxszz;
}


/* S Z Z  F R O M  S O Z */
/*----------------------------------------------------------------------------
	%%Function: SzzFromSoz

	String resources can't contain embedded '\0' characters, so we encode
	Szzs as Sozs.  An Soz contains a '\001' character for each desired
	embedded '\0'.  The resource is automatically terminated by a single
	'\0'.  This routine converts an Soz (as loaded from a resource) to an Szz.
	soz and szz can refer to the same location; the string will be converted
	in place.  szz will be one '\0' longer than soz, allow for this.
----------------------------------------------------------------------------*/
char FAR * SzzFromSoz(char FAR *soz, char FAR *szz)
{
	char FAR *szzSav = szz;

	while (*soz)
		{
		*szz++ = (*soz == '\001') ? '\0' : *soz;
		soz++;
		}
	*szz++ = '\0';
	*szz = '\0';

	return szzSav;
}


HGLOBAL StringToHGLOBAL (const char* str)
{
    char *p;
    HGLOBAL hMem = (HGLOBAL)NULL;
	
    if (str != NULL) {
        hMem = GlobalAlloc (GHND, strlen (str) + 1);
        p = (char*)GlobalLock (hMem);
        strcpy (p, str);
        GlobalUnlock (hMem);
    }
	
    return hMem;
}
