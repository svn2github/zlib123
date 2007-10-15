#include "stdafx.h"
#include "ClickHandler.h"

CClickHandler::CClickHandler(IHandleClick *pParent, int id)
{
	m_pParent = pParent;
	m_nId = id;
}


HRESULT CClickHandler::QueryInterface(REFIID riid, void  **ppvObject)
{
	//OutputDebugString(_T("CClickHandler::QueryInterface\n"));
	if (riid == IID_IUnknown) {
		//OutputDebugString(_T("CClickHandler::QueryInterface IUnknown\n"));
		*ppvObject = static_cast<IDispatch*>(this);
	} else if (riid == IID_IDispatch) {
		//OutputDebugString(_T("CClickHandler::QueryInterface IDispatch\n"));
		*ppvObject = static_cast<IDispatch*>(this);
	} else if (riid == __uuidof(_CommandBarButtonEvents)) {
		//OutputDebugString(_T("CClickHandler::QueryInterface _CommandBarButtonEvents\n"));
		*ppvObject = static_cast<IDispatch*>(this);
	} else {
		LPOLESTR iid;
		if (SUCCEEDED(StringFromIID(riid, &iid))) {
			//OutputDebugString(_T("Echec QI : "));
			//OutputDebugStringW(iid);
			CoTaskMemFree(iid);
			//OutputDebugString(_T("\n"));
		} else {
			//OutputDebugString(_T("Echec QI : ???\n"));
		}
		return E_NOINTERFACE;
	}
	return S_OK;
}

// No reference counting as livetime is handled by ConnectObject
ULONG CClickHandler::AddRef(void)
{
	return 2;
}

ULONG CClickHandler::Release(void)
{
	return 1;
}

HRESULT CClickHandler::GetTypeInfoCount(UINT *pctinfo)
{
	//OutputDebugString(_T("CClickHandler::GetTypeInfoCount\n"));
	return E_NOTIMPL;
}

HRESULT CClickHandler::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	//OutputDebugString(_T("CClickHandler::GetTypeInfo\n"));
	return E_NOTIMPL;
}

HRESULT CClickHandler::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	//OutputDebugString(_T("CClickHandler::GetIDsOfNames\n"));
	return E_NOTIMPL;
}

HRESULT CClickHandler::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	TCHAR sBuf[256];
	StringCchPrintf(sBuf, 256, _T("CClickHandler::Invoke : %d (%d)n"), dispIdMember, m_nId);
	OutputDebugString(sBuf);

	if (dispIdMember == 0x1) {
		return m_pParent->Clicked(m_nId);
	} else {
		return E_NOTIMPL;
	}
}
