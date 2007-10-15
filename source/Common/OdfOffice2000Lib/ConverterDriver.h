#pragma once

class CConverterDriver
{
public:
	CConverterDriver(void);
	virtual ~CConverterDriver(void);

	HRESULT CreateInstance();
	void Release();

	// Word conversion
	HRESULT OdtToDocx(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos);
	HRESULT DocxToOdt(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos);

	// Excel conversion
	HRESULT OdsToXlsx(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos);
	HRESULT XlsxToOds(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos);

	// Powerpoint conversion
	HRESULT OdpToPptx(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos);
	HRESULT PptxToOdp(BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos);

private:
	// Using IDispatch to driver converter because
	// 1. In out of proc, performance difference is negligible with dual interfaces
	// 2. It does not require TLB generation nor registration
	IDispatch *m_pConverter;

	// IDispatch invocation of all methods with the same signature
	HRESULT DoInvoke(DISPID dispid, BSTR bstrInputFile, BSTR bstrOutputFile, BOOL bShowUserInterface, WORD language, int centerPos);
};
