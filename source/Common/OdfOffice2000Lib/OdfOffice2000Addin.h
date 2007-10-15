#pragma once

#include "ClickHandler.h"
#include "ConverterDriver.h"

#define ID_IMPORT 1
#define ID_EXPORT 2

class OdfOffice2000Addin : public IHandleClick, public IMessageFilter {

public:
	// Construction / Destruction
	OdfOffice2000Addin(HMODULE hModule, UINT nDefaultTitle);
	virtual ~OdfOffice2000Addin();

private:
	DWORD m_dwImportAdvise;
	CClickHandler m_importHandler;

	DWORD m_dwExportAdvise;
	CClickHandler m_exportHandler;



	HMODULE m_hModule;
	UINT m_nDefaultTitle;

protected:
	Office2000::CommandBarControl *m_pImportButton;
	Office2000::CommandBarControl *m_pExportButton;

	WORD m_wLanguage;


	// Display Errors
	void ShowMessage(UINT idMsg, UINT idTitle);
	void ShowMessage(UINT idMsg);

	// Button Creation
	HRESULT CreateButton(CommandBarControls *pControls, UINT idResource, int idCommand, CommandBarControl **ppButton, CClickHandler *pClickHandler, DWORD *pdwAdvise);

	// Temporary filename generation
	BOOL GenerateTempName(LPTSTR sPath, DWORD dwPathLength, LPCTSTR sRoot, LPCTSTR sExtension);

	void LoadString(UINT uID, LPTSTR lpBuffer, int nBufferMax);


	// callback from ClickHandlers
	virtual HRESULT Clicked(int id) = 0;

	HRESULT ConfigureMenus(_CommandBars *pCommandBars, UINT uExportLabel, UINT uImportLabel, Office2000::CommandBarControl **ppOpenButton = NULL);
	HRESULT CleanMenus();

	CConverterDriver m_converter;


	IMessageFilter *m_pPreviousFilter;

	virtual DWORD STDMETHODCALLTYPE HandleInComingCall(DWORD dwCallType, HTASK htaskCaller, DWORD dwTickCount, LPINTERFACEINFO lpInterfaceInfo);
    virtual DWORD STDMETHODCALLTYPE RetryRejectedCall(HTASK htaskCallee, DWORD dwTickCount, DWORD dwRejectType);
    virtual DWORD STDMETHODCALLTYPE MessagePending(HTASK htaskCallee, DWORD dwTickCount, DWORD dwPendingType);


	void InstallMessageFilter();
	void UninstallMessageFilter();
};