/*----------------------------------------------------------------------------
	%%File: DBG.H

	Debugging support.
----------------------------------------------------------------------------*/

#ifndef DBG_H
#define DBG_H

#ifdef DEBUG

#define DeclareFileName		static char szAssertFileName [] = __FILE__;

// to reduce space usage in the 64K-maximum data segment, put debug-only
// string literals in the code segment.
#ifdef WIN16
#define CsConst(_type) _type __based(__segname("_CODE"))
#else
#define CsConst(_type) static _type const
#endif

#define Assert(_f) \
	{ \
	CsConst(char) _szFoo[] = "Assertion failed: "#_f;\
	if (!(_f)) \
		AssertFailed(szAssertFileName, __LINE__, _szFoo); \
	}
#define AssertSz(_f, _sz) \
	{ \
	if (!(_f)) \
		AssertFailed(szAssertFileName, __LINE__, _sz); \
	}
#define AssertDo    Assert

void AssertFailed(char *szFile, int li, const char FAR *szMessage);

#else

#define DeclareFileName

#define Assert(_f)
#define AssertSz(_f, _sz)
#define AssertDo(_f) (_f)

#endif

#endif // DBG_H
