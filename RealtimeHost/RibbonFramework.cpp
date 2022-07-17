#include "stdafx.h"
#include "RibbonFramework.h"
#include <Shobjidl.h>
#include <propvarutil.h>

RibbonEventLogger::RibbonEventLogger() {}

RibbonEventLogger::~RibbonEventLogger()
{
}

void STDMETHODCALLTYPE RibbonEventLogger::OnUIEvent(UI_EVENTPARAMS* pEventParams)
{
}

STDMETHODIMP_(ULONG) RibbonEventLogger::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) RibbonEventLogger::Release()
{
    LONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
        delete this;
    return cRef;
}

STDMETHODIMP RibbonEventLogger::QueryInterface(REFIID iid, void** ppv)
{
    if (!ppv)
        return E_POINTER;
    if (iid == __uuidof(IUnknown))
        *ppv = static_cast<IUnknown*>(this);
    else if (iid == __uuidof(IUIEventLogger))
        *ppv = static_cast<IUIEventLogger*>(this);
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

RibbonApplication::RibbonApplication(HWND hWindowFrame, CComPtr<RibbonCommandHandler> pCommandHandler) :m_hOwner(hWindowFrame), m_pCommandHandler(pCommandHandler) {}

RibbonApplication::~RibbonApplication()
{
}

STDMETHODIMP RibbonApplication::OnViewChanged(UINT32 nViewID, UI_VIEWTYPE typeID, IUnknown* pView, UI_VIEWVERB verb, INT32 uReasonCode)
{
    UNREFERENCED_PARAMETER(nViewID);
    UNREFERENCED_PARAMETER(typeID);
    UNREFERENCED_PARAMETER(pView);
    UNREFERENCED_PARAMETER(uReasonCode);

    HRESULT hr = E_FAIL;

    switch (verb)
    {
    case UI_VIEWVERB_SIZE:
        SendMessage(m_hOwner, UM_RIBBONVIEWCHANGED, 0, 0);
        hr = S_OK;
        break;
    }
    return hr;
}

STDMETHODIMP RibbonApplication::OnCreateUICommand(UINT32 nCmdID, UI_COMMANDTYPE typeID, IUICommandHandler** ppCommandHandler)
{
    UNREFERENCED_PARAMETER(typeID);
    HRESULT hr = S_OK;
    hr = m_pCommandHandler->QueryInterface(IID_PPV_ARGS(ppCommandHandler));
    return hr;
}

STDMETHODIMP RibbonApplication::OnDestroyUICommand(UINT32 commandId, UI_COMMANDTYPE typeID, IUICommandHandler* pCommandHandler)
{
    UNREFERENCED_PARAMETER(commandId);
    UNREFERENCED_PARAMETER(typeID);
    UNREFERENCED_PARAMETER(pCommandHandler);
    return E_NOTIMPL;
}

STDMETHODIMP_(ULONG) RibbonApplication::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) RibbonApplication::Release()
{
    LONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
        delete this;
    return cRef;
}

STDMETHODIMP RibbonApplication::QueryInterface(REFIID iid, void** ppv)
{
    if (!ppv)
        return E_POINTER;
    if (iid == __uuidof(IUnknown))
        *ppv = static_cast<IUnknown*>(this);
    else if (iid == __uuidof(IUIApplication))
        *ppv = static_cast<IUIApplication*>(this);
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

CComPtr<RibbonCommandHandler> RibbonApplication::GetCommandHandler()
{
    return m_pCommandHandler;
}

RibbonCommandHandler::RibbonCommandHandler() {}

RibbonCommandHandler::~RibbonCommandHandler()
{
}

HRESULT RibbonCommandHandler::Execute(UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY *key, const PROPVARIANT *pPropvarValue, IUISimplePropertySet *pCommandExecutionProperties)
{
	HRESULT hr = S_OK;
	if (m_executeListeners.count(nCmdID) > 0)
	{
        return m_executeListeners[nCmdID](nCmdID, verb, key, pPropvarValue, pCommandExecutionProperties);
	}
	return hr;
}

STDMETHODIMP RibbonCommandHandler::UpdateProperty(UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT *pPropvarCurrentValue, PROPVARIANT *pPropvarNewValue)
{
	HRESULT hr = E_NOTIMPL;
	if (m_updatePropertyListeners.count(nCmdID) > 0)
	{
        return m_updatePropertyListeners[nCmdID](nCmdID, key, pPropvarCurrentValue, pPropvarNewValue);
	}
	return hr;
}

STDMETHODIMP_(ULONG) RibbonCommandHandler::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) RibbonCommandHandler::Release()
{
    LONG cRef = InterlockedDecrement(&m_cRef);
    if (cRef == 0)
        delete this;
    return cRef;
}

STDMETHODIMP RibbonCommandHandler::QueryInterface(REFIID iid, void** ppv)
{
    if (!ppv)
        return E_POINTER;
    if (iid == __uuidof(IUnknown))
        *ppv = static_cast<IUnknown*>(this);
    else if (iid == __uuidof(IUICommandHandler))
        *ppv = static_cast<IUICommandHandler*>(this);
    else
    {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    AddRef();
    return S_OK;
}

void RibbonCommandHandler::AddOnUpdatePropertyListener(std::initializer_list<UINT> listCmd, OnUpdatePropertyListener handler)
{
    for (const auto& e : listCmd)
    {
        m_updatePropertyListeners.insert({ e, handler });
    }
}

RibbonFramework::RibbonFramework() {}

RibbonFramework::~RibbonFramework()
{
}

HRESULT RibbonFramework::InitializeFramework(HWND hWindowFrame, LPCWSTR pszRibbonResource)
{
    if (!hWindowFrame)
        return E_INVALIDARG;
    HRESULT hr = CoCreateInstance(CLSID_UIRibbonFramework, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pFramework));
    if (FAILED(hr))
        return hr;
    hr = CoCreateInstance(CLSID_UIRibbonImageFromBitmapFactory, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&m_pifbFactory));
    if (FAILED(hr))
        return hr;
    CComPtr<IUIEventingManager> pEventingManager;
    hr = m_pFramework->QueryInterface(IID_PPV_ARGS(&pEventingManager));
    if (FAILED(hr))
        return hr;
    hr = CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pWICFactory));
    if (FAILED(hr))
        return hr;
    m_pCommandHandler = new RibbonCommandHandler;
    m_pApplication = new RibbonApplication(hWindowFrame, m_pCommandHandler);
    hr = m_pFramework->Initialize(hWindowFrame, m_pApplication);
    if (FAILED(hr))
    {
        m_pApplication = nullptr;
        return hr;
    }
    hr = m_pFramework->LoadUI(GetModuleHandle(NULL), pszRibbonResource);
    if (FAILED(hr))
    {
        m_pApplication = nullptr;
        return hr;
    }
    m_pEventLogger = new RibbonEventLogger;
    hr = pEventingManager->SetEventLogger(m_pEventLogger);
    if (FAILED(hr))
    {
        m_pApplication = nullptr;
        return hr;
    }
    return hr;
}

HRESULT RibbonFramework::DestroyFramework()
{
    m_pWICFactory = nullptr;
    m_pifbFactory = nullptr;
    HRESULT hr = m_pFramework->Destroy();
    if (SUCCEEDED(hr))
        m_pFramework = nullptr;
    return hr;
}

BOOL RibbonFramework::GetRibbonHeight(UINT* puSize)
{
    if (m_pFramework)
    {
        CComPtr<IUIRibbon> pRibbon;
        HRESULT hr = m_pFramework->GetView(0, IID_PPV_ARGS(&pRibbon));
        if (SUCCEEDED(hr))
        {
            if (SUCCEEDED(pRibbon->GetHeight(puSize)))
                return TRUE;
            return FALSE;
        }
        return FALSE;
    }
    return FALSE;
}

CComPtr<RibbonCommandHandler> RibbonFramework::GetCommandHandler()
{
    return m_pCommandHandler;
}

CComPtr<IUIFramework> RibbonFramework::GetFramework()
{
    return m_pFramework;
}

//---------------------------------------------------------------------
// WICイメージからビットマップハンドルを生成
// 作成 20140913 KrK
// 引数 pBitmap：WICイメージ
//---------------------------------------------------------------------
HRESULT CreateHBitmapFromBitmapSource(IWICBitmapSource* pImage, HBITMAP& hBmp)
{
    HRESULT hr;                       // 処理結果
    WICPixelFormatGUID format;        // WICフォーマット
    BITMAPINFO bminfo = {};           // ビットマップヘッダ
    HDC hdc;                          // デバイスコンテキストハンドル
    UINT width = 0;                   // ビットマップ幅
    UINT height = 0;                  // ビットマップ高さ
    LPVOID imageBits = NULL;          // ビットマップイメージ

    // 32bppBGRAか確認
    hr = pImage->GetPixelFormat(&format);
    if (SUCCEEDED(hr))
    {
        if (format == GUID_WICPixelFormat32bppBGRA)
        {
            hr = S_OK;
        }
        else
        {
            hr = E_FAIL;
        }
    }
    // ビットマップサイズを取得
    if (SUCCEEDED(hr))
    {
        hr = pImage->GetSize(&width, &height);
    }
    // DIBを生成
    if (SUCCEEDED(hr))
    {
        bminfo.bmiHeader.biSize = sizeof(BITMAPINFO);
        bminfo.bmiHeader.biBitCount = 32;
        bminfo.bmiHeader.biCompression = BI_RGB;
        bminfo.bmiHeader.biWidth = width;
        bminfo.bmiHeader.biHeight = -static_cast<LONG>(height);
        bminfo.bmiHeader.biPlanes = 1;
        hdc = GetDC(HWND_DESKTOP);
        hBmp = CreateDIBSection(hdc, &bminfo, DIB_RGB_COLORS, &imageBits, NULL, 0);
        if (hBmp)
        {
            hr = S_OK;
        }
        else
        {
            hr = E_FAIL;
        }
        ReleaseDC(HWND_DESKTOP, hdc);
    }
    // イメージをコピー
    if (SUCCEEDED(hr))
    {
        hr = pImage->CopyPixels(NULL,
            width * 4,
            width * height * 4,
            reinterpret_cast<BYTE*>(imageBits));
    }
    else if (hBmp != NULL)
    {
        DeleteObject(hBmp);
        hBmp = NULL;
    }
    // ビットマップハンドルを返す
    return hr;
}

HRESULT RibbonFramework::GetImageFromBitmap(PCWSTR resourceName, PCWSTR resourceType, IUIImage** ppImg)
{
    HRSRC imageResHandle = FindResource(GetModuleHandle(NULL), resourceName, resourceType);
    HRESULT hr = imageResHandle ? S_OK : E_FAIL;

    HGLOBAL imageResDataHandle = NULL;
    if (SUCCEEDED(hr))
    {
        imageResDataHandle = LoadResource(GetModuleHandle(NULL), imageResHandle);
        hr = imageResDataHandle ? S_OK : E_FAIL;
    }

    void* pImageFile = NULL;
    if (SUCCEEDED(hr))
    {
        pImageFile = LockResource(imageResDataHandle);
        hr = pImageFile ? S_OK : E_FAIL;
    }
    DWORD imageFileSize = 0;
    if (SUCCEEDED(hr))
    {
        imageFileSize = SizeofResource(GetModuleHandle(NULL), imageResHandle);
        hr = imageFileSize ? S_OK : E_FAIL;
    }

    CComPtr<IWICStream> pStream;
    if (SUCCEEDED(hr))
        hr = m_pWICFactory->CreateStream(&pStream);

    if (SUCCEEDED(hr))
        hr = pStream->InitializeFromMemory(reinterpret_cast<BYTE*>(pImageFile), imageFileSize);

    CComPtr<IWICBitmapDecoder> pDecoder;
    if (SUCCEEDED(hr))
        hr = m_pWICFactory->CreateDecoderFromStream(pStream, NULL, WICDecodeMetadataCacheOnLoad, &pDecoder);

    CComPtr<IWICBitmapFrameDecode> pSource;
    if (SUCCEEDED(hr))
        hr = pDecoder->GetFrame(0, &pSource);

    CComPtr<IWICFormatConverter> pConverter;
    if (SUCCEEDED(hr))
    {
        // Convert the image format to 32bppPBGRA
        // (DXGI_FORMAT_B8G8R8A8_UNORM + D2D1_ALPHA_MODE_PREMULTIPLIED).
        hr = m_pWICFactory->CreateFormatConverter(&pConverter);
    }
    if (SUCCEEDED(hr))
        hr = pConverter->Initialize(pSource, GUID_WICPixelFormat32bppBGRA, WICBitmapDitherTypeNone, NULL, 0.f, WICBitmapPaletteTypeCustom);

    HBITMAP hGet = nullptr;
    if (SUCCEEDED(hr))
    {
        hr = CreateHBitmapFromBitmapSource(pConverter, hGet);
    }

    if (SUCCEEDED(hr))
    {
        hr = m_pifbFactory->CreateImage(hGet, UI_OWNERSHIP_TRANSFER, ppImg);
        if (FAILED(hr))
        {
            DeleteObject(hGet);
        }
    }
    return hr;
}

