#pragma once

static BOOL FRegisterClasses(HANDLE hkeyRoot, char *szSection,
			char *szzClasses, char *szzExts, char *szzNames, char *szModulePath);

FCE PASCAL Convert(HANDLE ghszFile, bool bExport);
FCE PASCAL Convert2(HANDLE ghszFile, bool bExport);