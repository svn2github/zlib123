// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include "OdfOffice2000Lib.h"
#include "Resource.h"

// CConnect
class ATL_NO_VTABLE CConnect : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<AddInDesignerObjects::_IDTExtensibility2, &AddInDesignerObjects::IID__IDTExtensibility2, &AddInDesignerObjects::LIBID_AddInDesignerObjects, 1, 0>,
	public OdfOffice2000Addin
{
public:
	CConnect();
	~CConnect();

DECLARE_REGISTRY_RESOURCEID(IDR_ADDIN)
DECLARE_NOT_AGGREGATABLE(CConnect)

BEGIN_COM_MAP(CConnect)
	COM_INTERFACE_ENTRY2(IDispatch, AddInDesignerObjects::IDTExtensibility2)
	COM_INTERFACE_ENTRY(AddInDesignerObjects::IDTExtensibility2)
	COM_INTERFACE_ENTRY(IMessageFilter)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:
	//IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, AddInDesignerObjects::ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(AddInDesignerObjects::ext_DisconnectMode RemoveMode, SAFEARRAY **custom );
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom );
	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom );
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom );

	// callback from ClickHandlers
	HRESULT Clicked(int id);

private:
	// Implementation fields
	_Application *m_pApplication;
	// Documents collection
	Documents *m_pDocuments;
	// Word 2007 Format number
	long _Word12SaveFormat;

	// Get Word 2007 Format number
	HRESULT GetWord2007SaveFormat();

	// Conversion from Odt to Word
	HRESULT OdtToDocx();
	// Conversion from Word to Odt
	HRESULT DocxToOdt();


	// Find top level window 
	HWND FindWord();
	// return the center of the window
	int ComputeCenter(HWND hWnd);


	HRESULT DeleteButton(CommandBar *pCommandBar, UINT idResource);
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
