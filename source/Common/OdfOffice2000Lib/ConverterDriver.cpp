#include "StdAfx.h"
#include "ConverterDriver.h"
#include "OdfOffice2000Lib.h"


// {94714634-B065-4f06-839E-5C03CADCB934}
static const GUID CLSID_ConverterWrapper =  { 0x94714634, 0xB065, 0x4f06, { 0x83, 0x9E, 0x5C, 0x03, 0xCA, 0xDC, 0xB9, 0x34 } };

#define DISPID_RELEASE 99

#define DISPID_ODTTODOCX 1
#define DISPID_DOCXTOODT 2
#define DISPID_ODSTOXLSX 3
#define DISPID_XLSXTOODS 4
#define DISPID_ODPTOPPTX 5
#define DISPID_PPTXTOODP 6

CConverterDriver::CConverterDriver(void) : m_pConverter(NULL)
{
}

CConverterDriver::~CConverterDriver(void)
{
	if (m_pConverter) {
		Release();
	}
}


HRESULT CConverterDriver::CreateInstance() {
	if (m_pConverter != NULL) {
		return S_OK;
	} else {
		HRESULT hr = CoCreateInstance(CLSID_ConverterWrapper, NULL, CLSCTX_ALL, IID_IDispatch, (void**)&m_pConverter);
		if (FAILED(hr)) {
			TraceError(hr, _T("CoCreateInstance(CLSID_ConverterWrapper)"));
		}
		return hr;
	}
}

void CConverterDriver::Release()
{
	if (m_pConverter != NULL) {
		// Call explicitely Exit
		DISPPARAMS params = { NULL, NULL, 0, 0 };
		VARIANT vtResult;
		VariantInit(&vtResult);
		EXCEPINFO excep;
		HRESULT hr = m_pConverter->Invoke(DISPID_RELEASE, IID_NULL, 0, DISPATCH_METHOD, &params, &vtResult, &excep, NULL);
		if (FAILED(hr)) {
#ifdef _DEBUG
			OutputDebugString(_T("Failed to invoke Exit\n"));
#endif
		}
		VariantClear(&vtResult);
		m_pConverter->Release();
		m_pConverter = NULL;
	}
}

// IDispatch invocation of all methods with the same signature
HRESULT CConverterDriver::DoInvoke(DISPID dispid, BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos)
{
	if (m_pConverter == NULL) {
		return E_FAIL;
	}

	// Array of parameters : reverse order when called by positions...
	VARIANT tabParams[5];
	tabParams[0].vt = VT_I4;
	tabParams[0].lVal = centerPos;
	tabParams[1].vt = VT_I2;
	tabParams[1].iVal = language;
	tabParams[2].vt = VT_BOOL;
	tabParams[2].boolVal = bShowUserInterface ? VARIANT_TRUE : VARIANT_FALSE;
	tabParams[3].vt = VT_BSTR;
	tabParams[3].bstrVal = SysAllocString(bstrOutputFile);
	tabParams[4].vt = VT_BSTR;
	tabParams[4].bstrVal = SysAllocString(bstrInputFile);

	DISPPARAMS params = { tabParams, NULL, sizeof(tabParams) / sizeof(tabParams[0]), 0 };
	VARIANT vtResult;
	VariantInit(&vtResult);

	EXCEPINFO excep;

	HRESULT hr = m_pConverter->Invoke(dispid, IID_NULL, 0, DISPATCH_METHOD, &params, &vtResult, &excep, NULL);
	VariantClear(&vtResult);
	return hr;
}


// Word conversion
HRESULT CConverterDriver::OdtToDocx(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos)
{
	return DoInvoke(DISPID_ODTTODOCX, bstrInputFile, bstrOutputFile, bShowUserInterface, language, centerPos);
}
HRESULT CConverterDriver::DocxToOdt(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos)
{
	return DoInvoke(DISPID_DOCXTOODT, bstrInputFile, bstrOutputFile, bShowUserInterface, language, centerPos);
}

// Excel conversion
HRESULT CConverterDriver::OdsToXlsx(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos)
{
	return DoInvoke(DISPID_ODSTOXLSX, bstrInputFile, bstrOutputFile, bShowUserInterface, language, centerPos);
}

HRESULT CConverterDriver::XlsxToOds(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos)
{
	return DoInvoke(DISPID_XLSXTOODS, bstrInputFile, bstrOutputFile, bShowUserInterface, language, centerPos);
}

// Powerpoint conversion
HRESULT CConverterDriver::OdpToPptx(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos)
{
	return DoInvoke(DISPID_ODPTOPPTX, bstrInputFile, bstrOutputFile, bShowUserInterface, language, centerPos);
}

HRESULT CConverterDriver::PptxToOdp(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos)
{
	return DoInvoke(DISPID_PPTXTOODP, bstrInputFile, bstrOutputFile, bShowUserInterface, language, centerPos);
}