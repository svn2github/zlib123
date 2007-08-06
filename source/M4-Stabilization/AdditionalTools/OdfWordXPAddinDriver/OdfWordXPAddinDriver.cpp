// OdfWordXPAddinDriver.cpp : Implementation of DLL Exports.


#include "stdafx.h"
#include "resource.h"
#include "OdfWordXPAddinDriver.h"


class COdfWordXPAddinDriverModule : public CAtlDllModuleT< COdfWordXPAddinDriverModule >
{
public :
	DECLARE_LIBID(LIBID_OdfWordXPAddinDriverLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_ODFWORDXPADDINDRIVER, "{B45472A9-CACC-4106-87D2-94E87472FB2B}")
};

COdfWordXPAddinDriverModule _AtlModule;


#ifdef _MANAGED
#pragma managed(push, off)
#endif

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	hInstance;
    return _AtlModule.DllMain(dwReason, lpReserved); 
}

#ifdef _MANAGED
#pragma managed(pop)
#endif




// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
    return _AtlModule.DllCanUnloadNow();
}


// {A77A8471-D758-457E-BD12-03E9FCBF25F4}
static const GUID CLSID_ODF = 
{ 0xA77A8471, 0xD758, 0x457E, { 0xBD, 0x12, 0x03, 0xE9, 0xFC, 0xBF, 0x25, 0xF4 } };

// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
#if TRUE
	//Don't instanciate the local object but the real ODF Add-in instead
	return CoGetClassObject(CLSID_ODF, CLSCTX_ALL, NULL, riid, ppv);
#else
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
#endif
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
	return hr;
}

