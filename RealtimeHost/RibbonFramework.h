#pragma once
#include "PropertySet.h"

#define UM_RIBBONVIEWCHANGED WM_USER + 0x17

class RibbonEventLogger : public IUIEventLogger
{
public:
	RibbonEventLogger();
	virtual ~RibbonEventLogger();
	// IUIEventLogger
	virtual void STDMETHODCALLTYPE OnUIEvent(UI_EVENTPARAMS* pEventParams) override;
	// IUnknown
	IFACEMETHODIMP_(ULONG) AddRef() override;
	IFACEMETHODIMP_(ULONG) Release() override;
	IFACEMETHODIMP QueryInterface(REFIID iid, void** ppv) override;

private:
	LONG m_cRef = 0;

};

using OnExecuteListener = std::function<HRESULT(UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties)>;
using OnUpdatePropertyListener = std::function<HRESULT(UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue)>;

class RibbonCommandHandler : public IUICommandHandler
{
public:
	RibbonCommandHandler();
	virtual ~RibbonCommandHandler();
	// IUICommandHandler
	STDMETHOD(Execute)(UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* ppropvarValue, IUISimplePropertySet* pCommandExecutionProperties) override;
	STDMETHOD(UpdateProperty)(UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* ppropvarCurrentValue, PROPVARIANT* ppropvarNewValue) override;
	// IUnknown
	IFACEMETHODIMP_(ULONG) AddRef() override;
	IFACEMETHODIMP_(ULONG) Release() override;
	IFACEMETHODIMP QueryInterface(REFIID iid, void** ppv) override;

	void AddOnExecuteListener(UINT cmdId, OnExecuteListener handler)
	{
		m_executeListeners.insert({ cmdId, handler });
	}
	void AddOnUpdatePropertyListener(UINT cmdId, OnUpdatePropertyListener handler)
	{
		m_updatePropertyListeners.insert({ cmdId, handler });
	}
	void AddOnUpdatePropertyListener(std::initializer_list<UINT> listCmd, OnUpdatePropertyListener handler);
private:
	LONG m_cRef = 0;
	std::map<UINT, OnExecuteListener> m_executeListeners;
	std::map<UINT, OnUpdatePropertyListener> m_updatePropertyListeners;
};

class RibbonApplication : public IUIApplication
{
public:
	RibbonApplication(HWND hWindowFrame, CComPtr<RibbonCommandHandler> pCommandHandler);
	virtual ~RibbonApplication();
	// IUIApplication
	STDMETHOD(OnViewChanged)(UINT32 nViewID, UI_VIEWTYPE typeID, IUnknown* pView, UI_VIEWVERB verb, INT32 uReasonCode) override;
	STDMETHOD(OnCreateUICommand)(UINT32 nCmdID, UI_COMMANDTYPE typeID, IUICommandHandler** ppCommandHandler) override;
	STDMETHOD(OnDestroyUICommand)(UINT32 commandId, UI_COMMANDTYPE typeID, IUICommandHandler* pCommandHandler) override;
	// IUnknown
	IFACEMETHODIMP_(ULONG) AddRef() override;
	IFACEMETHODIMP_(ULONG) Release() override;
	IFACEMETHODIMP QueryInterface(REFIID iid, void** ppv) override;

	CComPtr<RibbonCommandHandler> GetCommandHandler();
private:
	LONG m_cRef = 0;
	HWND m_hOwner = nullptr;
	CComPtr<RibbonCommandHandler> m_pCommandHandler;
};

class RibbonFramework
{
public:
	RibbonFramework();
	~RibbonFramework();

	HRESULT InitializeFramework(HWND hWindowFrame, LPCWSTR pszRibbonResource);
	HRESULT DestroyFramework();
	BOOL GetRibbonHeight(UINT* puSize);
	CComPtr<RibbonCommandHandler> GetCommandHandler();
	CComPtr<IUIFramework> GetFramework();
	HRESULT GetImageFromBitmap(PCWSTR resourceName, PCWSTR resourceType, IUIImage** ppImg);
private:
	CComPtr<IUIImageFromBitmap> m_pifbFactory;
	CComPtr<IUIFramework> m_pFramework;
	CComPtr<RibbonApplication> m_pApplication;
	CComPtr<RibbonCommandHandler> m_pCommandHandler;
	CComPtr<RibbonEventLogger> m_pEventLogger;
	CComPtr<IWICImagingFactory> m_pWICFactory;
};