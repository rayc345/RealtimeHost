#include "stdafx.h"
#include <shellapi.h>
#include <codecvt>
#include <sstream>
#include "MainFrm.h"
#include "resource.h"
#include "Utilities.h"

#if defined(_M_IX86)
#ifdef _DEBUG
#include "Debug/ids.h"
#else
#include "Release/ids.h"
#endif // _DEBUG
#elif defined(_M_AMD64)
#ifdef _DEBUG
#include "x64/Debug/ids.h"
#else
#include "x64/Release/ids.h"
#endif // _DEBUG
#endif

HANDLE hCom = nullptr;

// ERROR Code :
// 0 No error, operation succeeded
// 1 Failed to write
// 2 Failed to read
// 3 No data to read
// 5 Not enough buffer
uint8_t WriteBuffer(uint8_t uCommandID, const uint8_t* pWriteBuffer, uint16_t uWriteSize, uint8_t* pReadBuffer, uint16_t uReadSize, uint8_t& uWriteResult, uint16_t* uReadLength)
{
	uint8_t uBf[4];
	uBf[0] = (uint8_t)((uWriteSize + 1) >> 8);
	uBf[1] = (uint8_t)(uWriteSize + 1);
	DWORD dwCount = 0;
	BOOL bState = WriteFile(hCom, uBf, 2, &dwCount, NULL);
	if (!bState || dwCount != 2)
		return 1;
	uBf[0] = uCommandID;
	bState = WriteFile(hCom, uBf, 1, &dwCount, NULL);
	if (!bState || dwCount != 1)
		return 1;
	if (pWriteBuffer != nullptr && uWriteSize > 0)
	{
		bState = WriteFile(hCom, pWriteBuffer, uWriteSize, &dwCount, NULL);
		if (!bState || dwCount != uWriteSize)
			return 1;
	}
	bState = ReadFile(hCom, uBf, 3, &dwCount, NULL);
	if (!bState)
		return 2;
	if (dwCount != 3)
		return 3;
	uint16_t uReadResponse = (((uint16_t)uBf[0]) << 8) + (uint16_t)uBf[1];
	uWriteResult = uBf[2];
	if (uReadResponse >= 0)
	{
		if (pReadBuffer != nullptr && uReadLength != nullptr && uReadSize >= uReadResponse)
		{
			*uReadLength = uReadResponse;
			if (uReadResponse > 0)
			{
				bState = ReadFile(hCom, pReadBuffer, uReadResponse, &dwCount, NULL);
				if (!bState)
					return 2;
				if (dwCount != uReadResponse)
					return 3;
			}
		}
		else
		{
			DWORD dwDone = 0;
			while (dwDone < uReadResponse)
			{
				bState = ReadFile(hCom, uBf, min(uReadResponse - dwDone, 1), &dwCount, NULL);
				if (!bState)
					return 2;
				if (dwCount == 0)
					return 3;
				dwDone += dwCount;
			}
			if (pReadBuffer != nullptr && uReadLength != nullptr && uReadSize < uReadResponse)
				return 5;
		}
	}
	return 0;
}

decltype(auto) ReadTargetState = [=](int& uError, int& uState, int& uParams)
{
	uint8_t uReturnValue, uReadMessage[2];
	uint16_t uLenRead;
	uint8_t uHr = WriteBuffer(0x06, nullptr, 0, uReadMessage, 2, uReturnValue, &uLenRead);
	if (uHr != 0)
		uError = (int)uHr + 100;
	else
	{
		if (uLenRead != 2 || uReturnValue != 0)
			uError = (int)uReturnValue;
		else
		{
			uError = 0;
			uState = (int)uReadMessage[0];
			uParams = (int)uReadMessage[1];
		}
	}
};

decltype(auto) SetTargetState = [=](int& uError, uint8_t uState)
{
	uint8_t uReturnValue;
	uint8_t uHr = WriteBuffer(0x04, &uState, 1, nullptr, 0, uReturnValue, nullptr);
	if (uHr != 0)
		uError = (int)uHr + 100;
	else
	{
		if (uReturnValue != 0)
			uError = (int)uReturnValue;
		else
			uError = 0;
	}
};

decltype(auto) WriteParams = [=](int& uError, std::vector<std::tuple<uint16_t, float>>& vMsg, uint16_t& uSent)
{
	uint8_t uBuffer[8192];
	uint8_t uReadBuffer[4];
	uint8_t uReturnValue;
	std::tuple<uint16_t, float> qe;
	for (size_t si = 0; si < (uint16_t)vMsg.size(); si++)
	{
		qe = vMsg[si];
		uBuffer[0 + 6 * si] = (uint8_t)(std::get<0>(qe) >> 8);
		uBuffer[1 + 6 * si] = (uint8_t)std::get<0>(qe);
		float fValue = std::get<1>(qe);
		memcpy(uBuffer + 2 + 6 * si, &fValue, 4);
	}
	uint16_t uReadValue;
	uint8_t uHr = WriteBuffer(0x00, uBuffer, (uint16_t)vMsg.size() * 6, uReadBuffer, sizeof(uReadBuffer), uReturnValue, &uReadValue);
	if (uHr != 0)
		uError = (int)uHr + 100;
	else
	{
		if (uReturnValue != 0)
			uError = (int)uReturnValue;
		else
		{
			uError = 0;
			uSent = (((uint16_t)(*(uReadBuffer))) << 8) + (uint16_t)(*(uReadBuffer + 1));
			vMsg.erase(vMsg.begin(), vMsg.begin() + uSent);
		}
	}
};

decltype(auto) FirmwareUpload = [=](int& uError, uint8_t* pFirmware, uint32_t uSize)
{
	uint8_t uBuffer[2048];
	uint8_t uReturnValue;
	uint32_t uChunk = 2048;
	uint32_t uDone = 0;
	while (uDone <= uSize)
	{
		uint32_t uSizeThisTime;
		if (uSize - uDone < uChunk)
			uSizeThisTime = uSize - uDone;
		else
			uSizeThisTime = uChunk;

		if (uSizeThisTime != 0)
			memcpy(uBuffer, pFirmware + uDone, uSizeThisTime);
		uint8_t uHr = WriteBuffer(0x05, uBuffer, uSizeThisTime, nullptr, 0, uReturnValue, nullptr);
		if (uHr != 0)
			uError = (int)uHr + 100;
		else
		{
			if (uReturnValue != 0)
				uError = (int)uReturnValue;
			else
				uError = 0;
		}
		uDone += uSizeThisTime;
		if (uSizeThisTime == 0)
			break;
	}
};

static std::wstring string2wstring(std::string str)
{
	std::wstring result;
	//获取缓冲区大小，并申请空间，缓冲区大小按字符计算
	auto len = MultiByteToWideChar(CP_ACP, 0, str.c_str(), (int)str.size(), NULL, 0);
	TCHAR* buffer = new TCHAR[len + 1];
	//多字节编码转换成宽字节编码
	MultiByteToWideChar(CP_ACP, 0, str.c_str(), (int)str.size(), buffer, len);
	buffer[len] = '\0'; //添加字符串结尾
	//删除缓冲区并返回值
	result.append(buffer);
	delete[] buffer;
	return result;
}

decltype(auto) TextFetch = [=](int& uError, uint8_t* pBuffer, uint16_t uBufferSize, uint16_t& uReadBack)
{
	uint8_t uReturnValue;
	uint8_t uHr = WriteBuffer(0x01, nullptr, 0, pBuffer, uBufferSize, uReturnValue, &uReadBack);
	if (uHr != 0)
		uError = (int)uHr + 100;
	else
	{
		if (uReturnValue != 0)
			uError = (int)uReturnValue;
		else
			uError = 0;
	}
};

decltype(auto) SignalFetch = [=](int& uError, uint8_t* pBuffer, uint16_t uBufferSize, uint16_t& uReadBack)
{
	uint8_t uReturnValue;
	uint8_t uHr = WriteBuffer(0x02, nullptr, 0, pBuffer, uBufferSize, uReturnValue, &uReadBack);
	if (uHr != 0)
		uError = (int)uHr + 100;
	else
	{
		if (uReturnValue != 0)
			uError = (int)uReturnValue;
		else
			uError = 0;
	}
};

MainWindow::MainWindow()
{
}

MainWindow::~MainWindow()
{
}

void MainWindow::ReLayout()
{
	CRect rc;
	GetClientRect(&rc);
	UINT uHeight;
	if (m_RibboFramework.GetRibbonHeight(&uHeight))
	{
		rc.top = uHeight;
		if (m_hRichedit != nullptr)
			::SetWindowPos(m_hRichedit, HWND_BOTTOM, rc.left, rc.top + int(rc.Height() * 0.6f), rc.Width(), int(rc.Height() * 0.4f), SWP_SHOWWINDOW | SWP_NOZORDER);
		rc.bottom -= int(rc.Height() * 0.4f);
		if (m_hListView != nullptr)
			::SetWindowPos(m_hListView, HWND_TOP, rc.left, rc.top, int(rc.Width() * 0.6f), rc.Height(), SWP_SHOWWINDOW | SWP_NOZORDER);
		if (m_hListBox != nullptr)
			::SetWindowPos(m_hListBox, HWND_TOP, rc.left + int(rc.Width() * 0.6f), rc.top, int(rc.Width() * 0.4f), rc.Height(), SWP_SHOWWINDOW | SWP_NOZORDER);
	}
}

void MainWindow::ConnectionChanged()
{
	HRESULT hr;
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_CONNECT, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);

	PROPVARIANT Prop;
	hr = InitPropVariantFromBoolean(!m_bConnected, &Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_NEWPA, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_DELPA, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MOVEUP, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MOVEDOWN, UI_PKEY_Enabled, Prop);
	hr = PropVariantClear(&Prop);

	m_uTargetState = -1;
	m_cSignalsCount = -1;
	std::tuple<uint8_t*, uint16_t> e;
	while (m_DataStore.try_pop(e))
	{
		uint8_t* pArray = std::get<0>(e);
		delete[] pArray;
	}
	m_iReceiveFromTarget = 0;
	m_iTransmitToMATLAB = 0;

	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_RELOAD, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Enabled);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_PAUSEEXE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Enabled);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_PAUSEEXE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
}

void MainWindow::ParamSelChanged()
{
	HRESULT hr;
	PROPVARIANT Prop;
	hr = UIInitPropertyFromBoolean(UI_PKEY_Enabled, m_bParamSelected, &Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_VALUE, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MAXVALUE, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MINVALUE, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_STEP, UI_PKEY_Enabled, Prop);
	hr = PropVariantClear(&Prop);

	hr = UIInitPropertyFromBoolean(UI_PKEY_Enabled, m_bParamSelected && !m_bConnected, &Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_DELPA, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MOVEUP, UI_PKEY_Enabled, Prop);
	hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MOVEDOWN, UI_PKEY_Enabled, Prop);
	hr = PropVariantClear(&Prop);

	if (m_bParamSelected && m_uItemSelected == 0)
	{
		PROPVARIANT Prop;
		hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, FALSE, &Prop);
		hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MOVEUP, UI_PKEY_Enabled, Prop);
		hr = PropVariantClear(&Prop);
	}
	if (m_bParamSelected && m_uItemSelected == (m_vParams.size() - 1))
	{
		PROPVARIANT Prop;
		hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, FALSE, &Prop);
		hr = m_RibboFramework.GetFramework()->SetUICommandProperty(IDR_CMD_MOVEDOWN, UI_PKEY_Enabled, Prop);
		hr = PropVariantClear(&Prop);
	}

	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_MAXVALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SelectedItem);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_MINVALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SelectedItem);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_STEP, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SelectedItem);

	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_MaxValue);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_MinValue);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Increment);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_DecimalPlaces);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_DecimalValue);
}

void MainWindow::OnTargetStateChange()
{
	HRESULT hr;
	int error, state, param;
	ReadTargetState(error, state, param);
	if (error == 0)
	{
		m_uTargetState = state;
		if (m_uTargetState > 1)
			m_cSignalsCount = param;

		hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_RELOAD, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Enabled);
		hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_PAUSEEXE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Enabled);
		hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_PAUSEEXE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
	}
	else
	{
		if (error > 100)
			AddLogString(L"读取目标机状态失败,通信错误");
		else
			AddLogString(L"读取目标机状态失败,下位机返回错误");
	}
}

void MainWindow::AddLogString(std::wstring str)
{
	::SendMessage(m_hListBox, LB_ADDSTRING, (WPARAM)-1, (LPARAM)str.c_str());
	m_uLogCount++;
	::SendMessage(m_hListBox, LB_SETTOPINDEX, (WPARAM)(m_uLogCount - 1), 0);
}

void MainWindow::OnRefreshAllParam()
{
	std::vector<std::tuple<uint16_t, float>> vnewParams;
	for (size_t i = 0; i < m_vParams.size(); i++)
	{
		BOOL bFounded = FALSE;
		for (auto& e : m_vParamsToSend)
		{
			if (std::get<0>(e) == (uint16_t)i)
			{
				std::get<1>(e) = std::get<1>(m_vParams[i]);
				bFounded = TRUE;
			}
		}
		if (!bFounded)
			vnewParams.emplace_back(std::forward_as_tuple((uint16_t)i, std::get<1>(m_vParams[i])));
	}
	m_vParamsToSend.insert(m_vParamsToSend.end(), vnewParams.begin(), vnewParams.end());
}

BOOL MainWindow::RefreshComs()
{
	m_ComPorts.clear();
	HRESULT hres;
	hres = CoInitializeEx(0, COINIT_APARTMENTTHREADED);
	if (SUCCEEDED(hres) || hres == RPC_E_CHANGED_MODE)
	{
		hres = CoInitializeSecurity(
			NULL,
			-1,							 // COM authentication
			NULL,						 // Authentication services
			NULL,						 // Reserved
			RPC_C_AUTHN_LEVEL_DEFAULT,	 // Default authentication
			RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation
			NULL,						 // Authentication info
			EOAC_NONE,					 // Additional capabilities
			NULL						 // Reserved
		);

		if (SUCCEEDED(hres) || hres == RPC_E_TOO_LATE)
		{
			IWbemLocator* pLoc = NULL;

			hres = CoCreateInstance(
				CLSID_WbemLocator,
				0,
				CLSCTX_INPROC_SERVER,
				IID_IWbemLocator, (LPVOID*)&pLoc);

			if (SUCCEEDED(hres))
			{
				IWbemServices* pSvc = NULL;

				/*Connect to the root\cimv2 namespace with
					the current userand obtain pointer pSvc
					to make IWbemServices calls.*/
				hres = pLoc->ConnectServer(
					bstr_t(L"ROOT\\CIMV2"), // Object path of WMI namespace
					NULL,					// User name. NULL = current user
					NULL,					// User password. NULL = current
					0,						// Locale. NULL indicates current
					NULL,					// Security flags.
					0,						// Authority (for example, Kerberos)
					0,						// Context object
					&pSvc					// pointer to IWbemServices proxy
				);
				if (SUCCEEDED(hres))
				{
					hres = CoSetProxyBlanket(
						pSvc,						 // Indicates the proxy to set
						RPC_C_AUTHN_WINNT,			 // RPC_C_AUTHN_xxx
						RPC_C_AUTHZ_NONE,			 // RPC_C_AUTHZ_xxx
						NULL,						 // Server principal name
						RPC_C_AUTHN_LEVEL_CALL,		 // RPC_C_AUTHN_LEVEL_xxx
						RPC_C_IMP_LEVEL_IMPERSONATE, // RPC_C_IMP_LEVEL_xxx
						NULL,						 // client identity
						EOAC_NONE					 // proxy capabilities
					);
					if (SUCCEEDED(hres))
					{
						/*Use Win32_PnPEntity to find actual serial portsand USB - SerialPort devices
							This is done first, because it also finds some com0com devices, but names are worse*/
						CComPtr<IEnumWbemClassObject> pEnumerator;
						hres = pSvc->ExecQuery(
							bstr_t(L"WQL"),
							bstr_t(L"SELECT Name FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'"),
							WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
							NULL,
							&pEnumerator);

						if (SUCCEEDED(hres))
						{
							constexpr size_t max_ports = 30;
							IWbemClassObject* pclsObj[max_ports] = {};
							ULONG uReturn = 0;

							do
							{
								hres = pEnumerator->Next(WBEM_INFINITE, max_ports, pclsObj, &uReturn);
								if (SUCCEEDED(hres))
								{
									for (ULONG jj = 0; jj < uReturn; jj++)
									{
										VARIANT vtProp;
										HRESULT hr = pclsObj[jj]->Get(L"Name", 0, &vtProp, 0, 0);
										if (SUCCEEDED(hr))
										{
											/*Name should be for example "Serial Port for Barcode Scanner (COM13)"*/
											const std::wstring deviceName = vtProp.bstrVal;
											const std::wstring prefix = L"(COM";
											size_t ind = deviceName.find(prefix);
											if (ind != std::wstring::npos)
											{
												std::wstring nbr;
												for (size_t i = ind + prefix.length();
													i < deviceName.length() && isdigit(deviceName[i]); i++)
												{
													nbr += deviceName[i];
												}
												try
												{
													const int portNumber = _wtoi(nbr.c_str());
													m_ComPorts.emplace_back(deviceName);
												}
												catch (...)
												{
												}
											}
											VariantClear(&vtProp);
										}
										pclsObj[jj]->Release();
									}
								}
							} while (hres == WBEM_S_NO_ERROR);
						}
					}
					pSvc->Release();
				}
				pLoc->Release();
			}
		}
		CoUninitialize();
	}
	if (m_ComPorts.size() > 0)
	{
		decltype(auto) strSel = m_ComPorts[0];
		std::wregex e(LR"(COM\d+)");
		std::wsmatch m;
		bool found = std::regex_search(strSel, m, e);
		if (found && m.size() >= 1)
			m_ComSelected = m[0].str();
	}
	return SUCCEEDED(hres);
}

LONG_PTR lListviewProc;
LONG_PTR lEditProc;
HWND hMainwnd, hEdit;

LRESULT NewWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == WM_COMMAND)
		SendMessage(hMainwnd, message, wParam, lParam);
	return CallWindowProc((WNDPROC)lListviewProc, hWnd, message, wParam, lParam);
}

LRESULT NewEditWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == WM_KEYDOWN && VK_RETURN == wParam)
		SendMessage(hMainwnd, WM_COMMAND, wParam, (LPARAM)hEdit);
	return CallWindowProc((WNDPROC)lEditProc, hWnd, message, wParam, lParam);
}

int MainWindow::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	WSADATA wsaData;
	int err;
	err = WSAStartup(MAKEWORD(2, 2), &wsaData);
	if (err != 0) {
		return -1;
	}
	if (LOBYTE(wsaData.wVersion) != 2 || HIBYTE(wsaData.wVersion) != 2) {
		WSACleanup();
		return-1;
	}

	HRESULT hr;
	hr = m_RibboFramework.InitializeFramework(m_hWnd, L"APPLICATION_RIBBON");
	if (FAILED(hr))
		return -1;

	m_hListView = CreateWindowW(WC_LISTVIEW, L"", WS_CHILD | LVS_REPORT | LVS_EDITLABELS | LVS_SHOWSELALWAYS | LVS_SINGLESEL, 0, 0, 0, 0, m_hWnd, nullptr, GetModuleHandle(nullptr), nullptr);
	if (!::IsWindow(m_hListView))
		return -1;

	m_hRichedit = CreateWindowW(MSFTEDIT_CLASS, L"", ES_MULTILINE | ES_READONLY | WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOVSCROLL | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL, 0, 0, 0, 0, m_hWnd, nullptr, GetModuleHandle(nullptr), nullptr);
	if (!::IsWindow(m_hRichedit))
		return -1;

	lListviewProc = ::GetWindowLongPtr(m_hListView, GWLP_WNDPROC);
	::SetWindowLongPtr(m_hListView, GWLP_WNDPROC, (LONG_PTR)NewWndProc);

	m_hListBox = CreateWindowW(WC_LISTBOX, L"", WS_CHILD | WS_VSCROLL | LBS_NOTIFY | WS_BORDER | WS_VSCROLL | LBS_HASSTRINGS, 0, 0, 0, 0, m_hWnd, nullptr, GetModuleHandle(nullptr), nullptr);
	if (!::IsWindow(m_hListBox))
		return -1;

	CHARRANGE cr;
	cr.cpMin = -1;
	cr.cpMax = -1;

	::SendMessage(m_hRichedit, EM_EXSETSEL, 0, (LPARAM)&cr);

	PARAFORMAT2 pf;
	pf.cbSize = sizeof(PARAFORMAT2);
	pf.dwMask = PFM_LINESPACING;
	pf.dyLineSpacing = 230;
	pf.bLineSpacingRule = 4;

	::SendMessage(m_hRichedit, EM_SETPARAFORMAT, 0, LPARAM(&pf));

	decltype(auto) hFontOld = GetWindowFont(m_hListBox);
	LOGFONT lfText = {};
	SystemParametersInfoForDpi(SPI_GETICONTITLELOGFONT, sizeof(lfText), &lfText, FALSE, GetDpiForSystem());
	HFONT hFontNew = CreateFontIndirect(&lfText);
	if (hFontNew)
	{
		DeleteObject(hFontOld);
		EnumChildWindows(
			m_hWnd, [](HWND hWnd, LPARAM lParam) -> BOOL
			{
				::SendMessage(hWnd, WM_SETFONT, (WPARAM)lParam, MAKELPARAM(TRUE, 0));
				return TRUE;
			},
			(LPARAM)hFontNew);
	}

	LV_COLUMN lvColumn;
	TCHAR szString[][20] = { TEXT("编号"), TEXT("名称"), TEXT("值") };
	UINT iWidth[] = { 100 * GetDpiForSystem() / 96, 250 * GetDpiForSystem() / 96, 200 * GetDpiForSystem() / 96 };

	ListView_DeleteAllItems(m_hListView);

	lvColumn.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
	lvColumn.fmt = LVCFMT_LEFT;
	for (int i = 0; i < (sizeof(iWidth) / sizeof(UINT)); i++)
	{
		lvColumn.pszText = szString[i];
		lvColumn.cx = (int)iWidth[i];
		ListView_InsertColumn(m_hListView, i, &lvColumn);
	}
	ListView_SetExtendedListViewStyle(m_hListView, LVS_EX_FULLROWSELECT);
	return 0;
}

void MainWindow::OnSize(UINT nType, CSize size)
{
	ReLayout();
}

void MainWindow::OnDestroy()
{
	m_RibboFramework.DestroyFramework();
	PostQuitMessage(0);
	m_bThreadToStop = true;
	while (m_bThreadRunning);
	WSACleanup();
}

void MainWindow::OnClose()
{
	if (m_bChangeUnsaved)
	{
		auto rlt = MessageBox(m_hWnd, L"是否保存更改", L"保存", MB_YESNOCANCEL | MB_ICONQUESTION);
		if (rlt == IDYES)
		{
			if (m_strFileOpen.empty())
			{
				m_strFileOpen = funBrosweSaveLocation(this);
				if (!m_strFileOpen.empty())
					funSaveXML(this);
				else
				{
					SetMsgHandled(TRUE);
					return;
				}
			}
			else
				funSaveXML(this);
		}
		else if (rlt == IDCANCEL)
		{
			SetMsgHandled(TRUE);
			return;
		}
	}
	SetMsgHandled(FALSE);
}

void MainWindow::OnPaint(HDC dc)
{
	PAINTSTRUCT ps;
	dc = BeginPaint(m_hWnd, &ps);
	EndPaint(m_hWnd, &ps);
}

BOOL MainWindow::OnInitDialog(HWND wnd, LPARAM lInitParam)
{
	ReLayout();
	RefreshComs();
	decltype(auto) pComHandler = m_RibboFramework.GetCommandHandler();
	pComHandler->AddOnExecuteListener(IDR_CMD_NEW, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			if (m_bChangeUnsaved)
			{
				auto rlt = MessageBox(m_hWnd, L"是否保存更改", L"保存", MB_YESNOCANCEL | MB_ICONQUESTION);
				if (rlt == IDCANCEL)
					return S_OK;
				else if (rlt == IDYES)
				{
					if (m_strFileOpen.empty())
					{
						m_strFileOpen = funBrosweSaveLocation(this);
						if (!m_strFileOpen.empty())
							funSaveXML(this);
					}
					else
						funSaveXML(this);
				}
			}
			::SendMessage(m_hListView, LVM_DELETEALLITEMS, 0, 0);
			m_vParams.clear();
			m_strFileOpen = L"";
			m_bParamSelected = FALSE;
			m_bChangeUnsaved = FALSE;
			return S_OK;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_OPEN, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			if (m_bChangeUnsaved)
			{
				auto rlt = MessageBox(m_hWnd, L"是否保存更改", L"保存", MB_YESNOCANCEL | MB_ICONQUESTION);
				if (rlt == IDCANCEL)
					return S_OK;
				else if (rlt == IDYES)
				{
					if (m_strFileOpen.empty())
					{
						m_strFileOpen = funBrosweSaveLocation(this);
						if (!m_strFileOpen.empty())
							funSaveXML(this);
					}
					else
						funSaveXML(this);
				}
			}
			CComPtr<IFileDialog> pfd;
			DWORD dwFlags;
			HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
			hr = pfd->GetOptions(&dwFlags);
			hr = pfd->SetOptions(dwFlags | FOS_FORCEFILESYSTEM);
			COMDLG_FILTERSPEC fileType[] =
			{
				{L"XML files", L"*.xml"},
			};
			//	CComPtr<IShellItem> location;
			//	SHCreateItemFromParsingName(m_strConfigFolder.c_str(), nullptr, IID_PPV_ARGS(&location));
			//	pfd->SetFolder(location);
			//	pfd->AddPlace(location, FDAP_TOP);
			hr = pfd->SetFileTypes(ARRAYSIZE(fileType), fileType);
			hr = pfd->SetFileTypeIndex(1);
			hr = pfd->Show(m_hWnd);
			CComPtr<IShellItem> pShellItem;
			hr = pfd->GetResult(&pShellItem);
			std::wstring strFilePath;
			if (pShellItem)
			{
				LPWSTR filePath;
				hr = pShellItem->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &filePath);
				strFilePath = filePath;
				CoTaskMemFree(filePath);
				{
					pugi::xml_document doc;
					m_strFileOpen = strFilePath;
					doc.load_file(strFilePath.c_str());
					decltype(auto) rootNode = doc.child(L"Root");
					if (rootNode)
					{
						m_vParams.clear();
						for (decltype(auto) node = rootNode.first_child(); node; node = node.next_sibling())
						{
							float fValue = node.attribute(L"Value").as_float(0.f);
							float fMin = node.attribute(L"Min").as_float(0.f);
							float fMax = node.attribute(L"Max").as_float(1.f);
							fValue = funConstrain(fValue, fMin, fMax);
							m_vParams.emplace_back(std::forward_as_tuple(node.attribute(L"Name").as_string(), fValue, fMin, fMax,
								node.attribute(L"Step").as_float(1.f)));
						}
						m_bParamSelected = FALSE;
						m_bChangeUnsaved = FALSE;
						::SendMessage(m_hListView, LVM_DELETEALLITEMS, 0, 0);

						for (size_t i = 0; i < m_vParams.size(); ++i)
							funDrawItem(this, (int)i, TRUE);
						m_vParamsToSend.clear();
						OnRefreshAllParam();
					}
					else
						MessageBox(m_hWnd, L"文档无效", L"错误", MB_ICONERROR);
				}
			}
			return S_OK;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_SAVE, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			if (m_strFileOpen.empty())
			{
				m_strFileOpen = funBrosweSaveLocation(this);
				if (!m_strFileOpen.empty())
					funSaveXML(this);
			}
			else
				funSaveXML(this);
			return S_OK;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_SAVEAS, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			decltype(auto) sNewPath = funBrosweSaveLocation(this);
			if (!sNewPath.empty())
			{
				m_strFileOpen = sNewPath;
				funSaveXML(this);
			}
			return S_OK;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_EXIT, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			if (m_bChangeUnsaved)
			{
				auto rlt = MessageBox(m_hWnd, L"是否保存更改", L"保存", MB_YESNOCANCEL | MB_ICONQUESTION);
				if (rlt == IDCANCEL)
					return S_OK;
				else if (rlt == IDYES)
				{
					if (m_strFileOpen.empty())
					{
						m_strFileOpen = funBrosweSaveLocation(this);
						if (!m_strFileOpen.empty())
							funSaveXML(this);
						else
						{
							SetMsgHandled(TRUE);
							return S_OK;
						}
					}
					else
						funSaveXML(this);
				}
			}
			SendMessage(WM_CLOSE);
			return S_OK;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_HELP, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				ShellExecute(m_hWnd, L"open", LR"(https://github.com/rayc345/RealtimeHostDocument)", NULL, NULL, SW_SHOWNORMAL);
			}
			return S_OK;
		});
	// Ribbon Menu
	pComHandler->AddOnExecuteListener(IDR_CMD_CONNECT, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				if (pPropvarValue == nullptr)
				{
					if (m_bConnected)
					{
						KillTimer(0);
						int error;
						SetTargetState(error, 0x01);
						m_bConnected = FALSE;
						CloseHandle(hCom);
						hCom = nullptr;
						ConnectionChanged();
					}
					else
					{
						if (m_ComPorts.size() == 0 || m_ComSelected.empty())
							MessageBox(m_hWnd, L"不存在或未选择串口!", L"错误", MB_ICONERROR);
						else
						{
							std::wstring strOpen = m_ComSelected;
							std::wregex e(LR"((\d+))");
							std::wsmatch m;
							bool found = std::regex_search(m_ComSelected, m, e);
							int port = _wtoi(m[0].str().c_str());
							if (port > 9)
							{
								strOpen = LR"(\\.\)" + strOpen;
							}
							hCom = CreateFileW(strOpen.c_str(), GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
							if (hCom == INVALID_HANDLE_VALUE || hCom == nullptr)
							{
								WCHAR szErrorText[32];
								wsprintf(szErrorText, L"打开%s失败!", m_ComSelected.c_str());
								MessageBox(m_hWnd, szErrorText, L"错误", MB_ICONERROR);
							}
							else
							{
								SetupComm(hCom, 10240, 10240);
								COMMTIMEOUTS TimeOuts;
								TimeOuts.ReadIntervalTimeout = 20;
								TimeOuts.ReadTotalTimeoutMultiplier = 30;
								TimeOuts.ReadTotalTimeoutConstant = 200;
								TimeOuts.WriteTotalTimeoutMultiplier = 15;
								TimeOuts.WriteTotalTimeoutConstant = 200;
								SetCommTimeouts(hCom, &TimeOuts);
								DCB dcb;
								GetCommState(hCom, &dcb);
								dcb.BaudRate = 1500000;
								dcb.ByteSize = 8;
								dcb.Parity = NOPARITY;
								dcb.StopBits = ONESTOPBIT;
								SetCommState(hCom, &dcb);
								PurgeComm(hCom, PURGE_TXCLEAR | PURGE_RXCLEAR);
								DWORD dwErrorFlags;
								COMSTAT ComStat;
								ClearCommError(hCom, &dwErrorFlags, &ComStat);
								int state, param, error;
								ReadTargetState(error, state, param);
								if (error != 0)
								{
									m_bConnected = FALSE;
									CloseHandle(hCom);
									hCom = nullptr;
									MessageBox(m_hWnd, L"串口正常打开，但获取目标机状态错误。\n可能目标机未上电、未烧录引导器或串口通讯存在问题，\n若确认串口连接正常，可尝试复位目标机。", L"无法检测到目标机", MB_ICONWARNING);
								}
								else
								{
									m_bConnected = TRUE;
									ConnectionChanged();
									SetTimer(0, 200);
									OnRefreshAllParam();
								}
							}
						}
					}
					hr = S_OK;
				}
				else if (key && *key == UI_PKEY_SelectedItem)
				{
					if (m_bConnected != TRUE)
					{
						UINT selected;
						hr = UIPropertyToUInt32(*key, *pPropvarValue, &selected);
						decltype(auto) strSel = m_ComPorts[selected];
						std::wregex e(LR"(COM\d+)");
						std::wsmatch m;
						bool found = std::regex_search(strSel, m, e);
						if (found && m.size() >= 1)
							m_ComSelected = m[0].str();
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_CONNECT, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SelectedItem);
					}
				}
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_REFRESH, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			RefreshComs();
			HRESULT hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_CONNECT, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_ItemsSource);
			return S_OK;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_LISTEN, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			if (m_sServer == INVALID_SOCKET)
			{
				do
				{
					SOCKADDR_IN ServerAddr = { 0 };
					ServerAddr.sin_family = AF_INET;
					ServerAddr.sin_port = htons(m_hListenPort);
					auto rlt = inet_pton(AF_INET, "127.0.0.1", &ServerAddr.sin_addr.S_un.S_addr);
					if (rlt != 1)
						break;

					m_sServer = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
					if (m_sServer == INVALID_SOCKET)
						break;
					m_hServerEvent = WSACreateEvent();
					WSAEventSelect(m_sServer, m_hServerEvent, FD_ACCEPT);

					if (SOCKET_ERROR != bind(m_sServer, (SOCKADDR*)&ServerAddr, sizeof(ServerAddr)) && SOCKET_ERROR != listen(m_sServer, SOMAXCONN))
					{
						m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_LISTEN, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_BooleanValue);
						m_bThreadToStop = false;
						m_bThreadRunning = false;
						std::thread([&]() {
							m_bThreadRunning = true;
							while (!m_bThreadToStop)
							{
								WSANETWORKEVENTS networkEvents;          //网络事件结构
								if (iTotal < WSA_MAXIMUM_WAIT_EVENTS) //这个值是64
								{
									if (0 == WSAEnumNetworkEvents(m_sServer, m_hServerEvent, &networkEvents))
									{
										if (networkEvents.lNetworkEvents & FD_ACCEPT) //如果等于FD_ACCEPT，相与就为1
										{
											if (0 == networkEvents.iErrorCode[FD_ACCEPT_BIT]) //检查有无网络错误
											{
												//接受请求
												SOCKADDR_IN addrServer = { 0 };
												int iaddrLen = sizeof(addrServer);
												m_sClient[iTotal] = accept(m_sServer, (SOCKADDR*)&addrServer, &iaddrLen);
												//为新的client注册网络事件
												m_hClientEvent[iTotal] = WSACreateEvent();
												WSAEventSelect(m_sClient[iTotal], m_hClientEvent[iTotal], FD_READ | FD_WRITE | FD_CLOSE);
												iTotal++;
											}
											else //错误处理
											{
												int iError = networkEvents.iErrorCode[FD_ACCEPT_BIT];
											}
										}
									}
								}
								if (iTotal > 0)
								{
									//等待网络事件
									DWORD dwIndex = WSAWaitForMultipleEvents(iTotal, m_hClientEvent, FALSE, 100, FALSE);
									//当前的事件对象
									WSAEVENT curEvent = m_hClientEvent[dwIndex];
									//当前的套接字
									SOCKET sCur = m_sClient[dwIndex];
									//网络事件结构
									WSANETWORKEVENTS networkEvents;
									if (0 == WSAEnumNetworkEvents(sCur, curEvent, &networkEvents))
									{
										if (networkEvents.lNetworkEvents & FD_READ)
										{
											if (0 == networkEvents.iErrorCode[FD_READ_BIT])
											{
												char buf[4];
												int iRet = recv(sCur, buf, 1, 0);
												if (iRet != SOCKET_ERROR)
												{
													if (buf[0] == 0)
													{
														send(sCur, (char*)&m_cSignalsCount, 1, 0);
													}
													else if (buf[0] == 1)
													{
														std::tuple<uint8_t*, uint16_t> e;
														if (m_DataStore.try_pop(e))
														{
															uint16_t uArray = std::get<1>(e);
															uint8_t* pArray = std::get<0>(e);

															uint16_t uPairNub = uArray / 4 / m_cSignalsCount;
															send(sCur, (char*)&uPairNub, 2, 0);
															send(sCur, (char*)pArray, uArray, 0);
															delete[] pArray;

															m_iTransmitToMATLAB += uPairNub;

															std::wostringstream stm;
															stm << L"向MATLAB累计发送" << m_iTransmitToMATLAB << L"组信号";
															AddLogString(stm.str());

														}
														else
														{
															uint16_t uPairNub = 0;
															send(sCur, (char*)&uPairNub, 2, 0);
														}
													}
												}
											}
											else //错误处理
											{
												int iError = networkEvents.iErrorCode[FD_ACCEPT_BIT];
												break;
											}
										}
										else if (networkEvents.lNetworkEvents & FD_CLOSE) //client关闭
										{
											iTotal--;
										}
									}
								}
								else
									Sleep(100);
							}
							m_bThreadRunning = false;
							AddLogString(L"监听停止。");
							}).detach();
					}
					else
					{
						closesocket(m_sServer);
						break;
					}
					return S_OK;
				} while (0);
				closesocket(m_sServer);
				m_sServer = INVALID_SOCKET;
				MessageBox(m_hWnd, L"监听端口失败", L"错误", MB_ICONERROR);
			}
			else
			{
				m_bThreadToStop = true;
				while (m_bThreadRunning);
				for (int i = 0; i < iTotal; i++)
					closesocket(m_sClient[i]);
				closesocket(m_sServer);
				m_sServer = INVALID_SOCKET;
				iTotal = 0;
			}
			return S_OK;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_FIRM, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected;
					hr = UIPropertyToUInt32(*key, *pPropvarValue, &selected);
					m_uFirmSel = selected;
					hr = S_OK;
				}
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_FIRMBROW, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				CComPtr<IFileDialog> pfd;
				DWORD dwFlags;
				HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
				hr = pfd->GetOptions(&dwFlags);
				hr = pfd->SetOptions(dwFlags | FOS_FORCEFILESYSTEM);
				COMDLG_FILTERSPEC fileType[] =
				{
					{L"HEX files", L"*.hex"},
					{L"Binary files", L"*.bin"},
					{L"All files", L"*.bin;*.hex"},
				};
				hr = pfd->SetFileTypes(ARRAYSIZE(fileType), fileType);
				hr = pfd->SetFileTypeIndex(3);
				hr = pfd->Show(m_hWnd);
				CComPtr<IShellItem> pShellItem;
				hr = pfd->GetResult(&pShellItem);
				LPWSTR filePath;
				std::wstring strFilePath;
				if (pShellItem)
				{
					hr = pShellItem->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &filePath);
					strFilePath = filePath;
					CoTaskMemFree(filePath);
					for (decltype(auto) e : m_FirmPaths)
					{
						if (e == strFilePath)
							return hr;
					}
					m_FirmPaths.emplace_back(strFilePath);
					hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_FIRM, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_ItemsSource);
				}
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_PAUSEEXE, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				BOOL bSelected;
				if (*key == UI_PKEY_BooleanValue)
				{
					hr = UIPropertyToBoolean(*key, *pPropvarValue, &bSelected);
					uint8_t uState;
					if (bSelected)
						uState = 0x01;
					else
						uState = 0x02;
					int error;
					SetTargetState(error, uState);
					if (error != 0)
					{
						AddLogString(L"设置目标机状态失败。");
					}
					else
					{
						if (bSelected)
							AddLogString(L"成功暂停目标机。");
						else
							AddLogString(L"成功恢复目标机。");
					}
					OnTargetStateChange();
				}
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_RELOAD, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				int error;
				SetTargetState(error, 0x00);
				if (error != 0)
				{
					AddLogString(L"设置目标机状态失败。");
				}
				else
				{
					AddLogString(L"成功复位目标机。");
				}
				OnTargetStateChange();
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_NEWPA, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				UINT uInserPos = 0;
				auto newItem = std::forward_as_tuple(L"新属性", 1.f, 0.f, 1.f, 1.f);
				if (m_bParamSelected)
				{
					m_vParams.insert(m_vParams.begin() + m_uItemSelected + 1, newItem);

					funDrawItem(this, (int)m_vParams.size() - 1, TRUE);
					for (size_t i = m_uItemSelected; i < m_vParams.size(); ++i)
						funDrawItem(this, (int)i);
					ListView_SetItemState(m_hListView, m_uItemSelected, ~(LVIS_SELECTED | LVIS_FOCUSED), LVIS_SELECTED | LVIS_FOCUSED);
					ListView_SetItemState(m_hListView, m_uItemSelected + 1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
				}
				else
				{
					m_vParams.insert(m_vParams.end(), newItem);
					funDrawItem(this, (int)m_vParams.size() - 1, TRUE);
					ListView_SetItemState(m_hListView, (int)m_vParams.size() - 1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
					m_bParamSelected = TRUE;
					m_uItemSelected = (int)m_vParams.size() - 1;
				}
				m_bChangeUnsaved = TRUE;
				ParamSelChanged();
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_DELPA, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				::SendMessage(m_hListView, LVM_DELETEITEM, m_uItemSelected, 0);
				m_vParams.erase(m_vParams.begin() + m_uItemSelected);
				for (size_t i = m_uItemSelected; i < m_vParams.size(); ++i)
					funDrawItem(this, (int)i);
				m_bChangeUnsaved = TRUE;
				m_bParamSelected = FALSE;
				ParamSelChanged();
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_VALUE, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				DECIMAL dec;
				hr = UIPropertyToDecimal(*key, *pPropvarValue, &dec);
				float fNew;
				hr = VarR4FromDec(&dec, &fNew);
				if (FAILED(hr))
					return hr;
				std::get<1>(m_vParams[m_uItemSelected]) = fNew;
				bool bfound = false;
				std::for_each(m_vParamsToSend.begin(), m_vParamsToSend.end(), [=, &bfound](auto& value)
					{
						if (std::get<0>(value) == m_uItemSelected)
							std::get<1>(value) = fNew;
						bfound = true;
					});
				if (!bfound)
					m_vParamsToSend.emplace_back(std::forward_as_tuple((uint16_t)m_uItemSelected, fNew));
				funDrawItem(this, m_uItemSelected);
				m_bChangeUnsaved = TRUE;
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_MOVEUP, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				auto tmp = m_vParams[m_uItemSelected];
				m_vParams[m_uItemSelected] = m_vParams[m_uItemSelected - 1];
				m_vParams[m_uItemSelected - 1] = std::move(tmp);
				funDrawItem(this, m_uItemSelected - 1);
				funDrawItem(this, m_uItemSelected);
				m_bChangeUnsaved = TRUE;
				ListView_SetItemState(m_hListView, m_uItemSelected, ~(LVIS_SELECTED | LVIS_FOCUSED), LVIS_SELECTED | LVIS_FOCUSED);
				ListView_SetItemState(m_hListView, m_uItemSelected - 1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
				ParamSelChanged();
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_MOVEDOWN, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				auto tmp = m_vParams[m_uItemSelected];
				m_vParams[m_uItemSelected] = m_vParams[m_uItemSelected + 1];
				m_vParams[m_uItemSelected + 1] = std::move(tmp);
				funDrawItem(this, m_uItemSelected + 1);
				funDrawItem(this, m_uItemSelected);
				m_bChangeUnsaved = TRUE;
				ListView_SetItemState(m_hListView, m_uItemSelected, ~(LVIS_SELECTED | LVIS_FOCUSED), LVIS_SELECTED | LVIS_FOCUSED);
				ListView_SetItemState(m_hListView, m_uItemSelected + 1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
				ParamSelChanged();
			}
			return hr;
		});
	std::vector<float> MaxList = { 1000.f, 100.f, 10.f, 1.f, 0.1f, 0.01f, 0.001f };
	std::vector<float> MinList = { 1000.f, 100.f, 10.f, 1.f, 0.f, 0.1f, 0.01f, 0.001f };
	std::vector<float> StepList = { 1000.f, 100.f, 10.f, 1.f, 0.1f, 0.01f, 0.001f };
	pComHandler->AddOnExecuteListener(IDR_CMD_MAXVALUE, [&, MaxList](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected;
					float fMax = 0;
					hr = UIPropertyToUInt32(*key, *pPropvarValue, &selected);
					if (selected == UI_COLLECTION_INVALIDINDEX)
					{
						if (pCommandExecutionProperties != nullptr)
						{
							PROPVARIANT var;
							if (SUCCEEDED(pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var)))
							{
								BSTR bstr = var.bstrVal;
								fMax = wcstof(bstr, nullptr);
								PropVariantClear(&var);
							}
						}
					}
					else
						fMax = *(MaxList.begin() + selected);
					if (fMax >= std::get<2>(m_vParams[m_uItemSelected]))
					{
						std::get<1>(m_vParams[m_uItemSelected]) = funConstrain(std::get<1>(m_vParams[m_uItemSelected]), std::get<2>(m_vParams[m_uItemSelected]), fMax);
						std::get<3>(m_vParams[m_uItemSelected]) = fMax;
						ParamSelChanged();
						funDrawItem(this, m_uItemSelected);
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_MaxValue);
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_DecimalValue);
					}
					m_bChangeUnsaved = TRUE;
				}
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_MINVALUE, [&, MinList](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected;
					float fMin = 0;
					hr = UIPropertyToUInt32(*key, *pPropvarValue, &selected);
					if (selected == UI_COLLECTION_INVALIDINDEX)
					{
						if (pCommandExecutionProperties != nullptr)
						{
							PROPVARIANT var;
							if (SUCCEEDED(pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var)))
							{
								BSTR bstr = var.bstrVal;
								fMin = wcstof(bstr, nullptr);
								PropVariantClear(&var);
							}
						}
					}
					else
						fMin = *(MinList.begin() + selected);
					if (fMin <= std::get<3>(m_vParams[m_uItemSelected]))
					{
						std::get<1>(m_vParams[m_uItemSelected]) = funConstrain(std::get<1>(m_vParams[m_uItemSelected]), fMin, std::get<3>(m_vParams[m_uItemSelected]));
						std::get<2>(m_vParams[m_uItemSelected]) = fMin;
						ParamSelChanged();
						funDrawItem(this, m_uItemSelected);
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_MinValue);
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_DecimalValue);
					}
					m_bChangeUnsaved = TRUE;
				}
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_STEP, [&, StepList](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				if (key && *key == UI_PKEY_SelectedItem)
				{
					UINT selected;
					hr = UIPropertyToUInt32(*key, *pPropvarValue, &selected);
					if (selected == UI_COLLECTION_INVALIDINDEX)
					{
						if (pCommandExecutionProperties != nullptr)
						{
							PROPVARIANT var;
							if (SUCCEEDED(pCommandExecutionProperties->GetValue(UI_PKEY_Label, &var)))
							{
								BSTR bstr = var.bstrVal;
								float fNew = wcstof(bstr, nullptr);
								PropVariantClear(&var);
								std::get<4>(m_vParams[m_uItemSelected]) = fNew;
							}
						}
					}
					else
						std::get<4>(m_vParams[m_uItemSelected]) = *(StepList.begin() + selected);
					ParamSelChanged();
					hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_VALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_Increment);
					m_bChangeUnsaved = TRUE;
				}
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_CLEARLOG, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				::SendMessage(m_hListBox, LB_RESETCONTENT, 0, 0);
				m_uLogCount = 0;
			}
			return hr;
		});
	pComHandler->AddOnExecuteListener(IDR_CMD_CLEARTXT, [&](UINT nCmdID, UI_EXECUTIONVERB verb, const PROPERTYKEY* key, const PROPVARIANT* pPropvarValue, IUISimplePropertySet* pCommandExecutionProperties) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			UNREFERENCED_PARAMETER(pCommandExecutionProperties);
			HRESULT hr = E_FAIL;
			if (verb == UI_EXECUTIONVERB_EXECUTE)
			{
				::SetWindowText(m_hRichedit, L"");
			}
			return hr;
		});
	//update property
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_CONNECT, [&](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_Categories)
			{
				CComPtr<IUICollection> pCollection;
				hr = pPropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr))
					return hr;

				CComPtr<CPropertySet> pCAN;
				hr = CPropertySet::CreateInstance(&pCAN);
				if (FAILED(hr))
					return hr;

				WCHAR wszLabel[MAX_RESOURCE_LENGTH];
				swprintf_s(wszLabel, MAX_RESOURCE_LENGTH, L"串口CAN桥接");
				pCAN->InitializeCategoryProperties(wszLabel, 0);
				pCollection->Add(pCAN);

				CComPtr<CPropertySet> pUART;
				hr = CPropertySet::CreateInstance(&pUART);
				if (FAILED(hr))
					return hr;
				swprintf_s(wszLabel, MAX_RESOURCE_LENGTH, L"串口");
				pUART->InitializeCategoryProperties(wszLabel, 1);
				pCollection->Add(pUART);

				hr = S_OK;
			}
			else if (key == UI_PKEY_ItemsSource)
			{
				CComPtr<IUICollection> pCollection;
				hr = pPropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				hr = pCollection->Clear();
				for (size_t i = 0; i < m_ComPorts.size(); ++i)
				{
					CComPtr<CPropertySet> pItem;
					hr = CPropertySet::CreateInstance(&pItem);

					CComPtr<IUIImage> pImg;
					decltype(auto) strPortname = m_ComPorts[i];
					//if (std::regex_match(strPortname, std::wregex(LR"(.+)" + m_ComSelected + LR"(.+)")))
					//	hr = m_RibboFramework.GetImageFromBitmap(MAKEINTRESOURCE(IDP_PNG_CHECKMARK), L"PNG", &pImg);
					//else
					//	hr = m_RibboFramework.GetImageFromBitmap(MAKEINTRESOURCE(IDP_PNG_BLANK), L"PNG", &pImg);
					pItem->InitializeItemProperties(pImg, strPortname.c_str(), 1);
					hr = pCollection->Add(pItem);
				}
				hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_CONNECT, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SelectedItem);
				hr = S_OK;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				UINT uSel = 0;
				if (!m_ComSelected.empty())
				{
					for (size_t i = 0; i < m_ComPorts.size(); i++)
					{
						if (std::regex_match(m_ComPorts[i], std::wregex(LR"(.+)" + m_ComSelected + LR"(.+)")))
							uSel = (UINT)i;
					}
				}
				hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, uSel, pPropvarNewValue);
			}
			else if (key == UI_PKEY_BooleanValue)
				hr = UIInitPropertyFromBoolean(UI_PKEY_BooleanValue, m_bConnected, pPropvarNewValue);
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_LISTEN, [&](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_BooleanValue)
			{
				hr = InitPropVariantFromBoolean(m_sServer != INVALID_SOCKET, pPropvarNewValue);
			}
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_FIRM, [&](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_Categories)
				hr = S_FALSE;
			else if (key == UI_PKEY_ItemsSource)
			{
				CComPtr<IUICollection> pCollection;
				hr = pPropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				hr = pCollection->Clear();
				for (size_t i = 0; i < m_FirmPaths.size(); ++i)
				{
					CComPtr<CPropertySet> pItem;
					hr = CPropertySet::CreateInstance(&pItem);

					CComPtr<IUIImage> pImg;
					//if (m_uFirmSel == i)
					//	hr = m_RibboFramework.GetImageFromBitmap(MAKEINTRESOURCE(IDP_PNG_CHECKMARK), L"PNG", &pImg);
					//else
					//	hr = m_RibboFramework.GetImageFromBitmap(MAKEINTRESOURCE(IDP_PNG_BLANK), L"PNG", &pImg);
					pItem->InitializeItemProperties(pImg, m_FirmPaths[i].c_str(), UI_COLLECTION_INVALIDINDEX);
					hr = pCollection->Add(pItem);
				}
				hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_FIRM, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_SelectedItem);
				hr = S_OK;
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				if (m_FirmPaths.size() > 0)
					hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, m_uFirmSel, pPropvarNewValue);
			}
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_PAUSEEXE, [&](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_Enabled)
			{
				hr = InitPropVariantFromBoolean(m_uTargetState > 0, pPropvarNewValue);
			}
			else if (key == UI_PKEY_BooleanValue)
			{
				hr = InitPropVariantFromBoolean(m_uTargetState == 1, pPropvarNewValue);
			}
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_RELOAD, [&](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_Enabled)
			{
				hr = InitPropVariantFromBoolean(m_uTargetState > 0, pPropvarNewValue);
			}
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_VALUE, [&](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_MaxValue)
			{
				if (m_bParamSelected)
				{
					float fValue = std::get<3>(m_vParams[m_uItemSelected]);
					DECIMAL dec;
					hr = VarDecFromR4(fValue, &dec);
					if (FAILED(hr))
						return hr;
					hr = UIInitPropertyFromDecimal(UI_PKEY_MaxValue, dec, pPropvarNewValue);
				}
			}
			else if (key == UI_PKEY_MinValue)
			{
				if (m_bParamSelected)
				{
					float fValue = std::get<2>(m_vParams[m_uItemSelected]);
					DECIMAL dec;
					hr = VarDecFromR4(fValue, &dec);
					if (FAILED(hr))
						return hr;
					hr = UIInitPropertyFromDecimal(UI_PKEY_MinValue, dec, pPropvarNewValue);
				}
			}
			else if (key == UI_PKEY_Increment)
			{
				if (m_bParamSelected)
				{
					float fValue = std::get<4>(m_vParams[m_uItemSelected]);
					DECIMAL dec;
					hr = VarDecFromR4(fValue, &dec);
					if (FAILED(hr))
						return hr;
					hr = UIInitPropertyFromDecimal(UI_PKEY_Increment, dec, pPropvarNewValue);
				}
			}
			else if (key == UI_PKEY_DecimalValue)
			{
				if (m_bParamSelected)
				{
					float fCurrentValue = std::get<1>(m_vParams[m_uItemSelected]);
					float fNewValue = funConstrain(std::get<1>(m_vParams[m_uItemSelected]), std::get<2>(m_vParams[m_uItemSelected]), std::get<3>(m_vParams[m_uItemSelected]));
					if (fNewValue != fCurrentValue)
					{
						std::get<1>(m_vParams[m_uItemSelected]) = fNewValue;
						m_vParamsToSend.emplace_back(std::forward_as_tuple((uint16_t)m_uItemSelected, fNewValue));
						funDrawItem(this, m_uItemSelected);
					}
					DECIMAL dec;
					hr = VarDecFromR4(fNewValue, &dec);
					if (FAILED(hr))
						return hr;
					hr = UIInitPropertyFromDecimal(UI_PKEY_DecimalValue, dec, pPropvarNewValue);
				}
			}
			else if (key == UI_PKEY_DecimalPlaces)
			{
				UINT value = 3;
				hr = UIInitPropertyFromUInt32(UI_PKEY_DecimalPlaces, value, pPropvarNewValue);
			}
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_MAXVALUE, [&, MaxList](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_Categories)
				hr = S_FALSE;
			else if (key == UI_PKEY_ItemsSource)
			{
				CComPtr<IUICollection> pCollection;
				hr = pPropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr))
					return hr;
				hr = pCollection->Clear();
				for (size_t i = 0; i < MaxList.size(); i++)
				{
					CComPtr<CPropertySet> pItem;
					hr = CPropertySet::CreateInstance(&pItem);
					if (FAILED(hr))
						return hr;
					pItem->InitializeItemProperties(NULL, funConvert(*(MaxList.begin() + i)).c_str(), UI_COLLECTION_INVALIDINDEX);
					hr = pCollection->Add(pItem);
					hr = S_OK;
				}
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				if (m_bParamSelected)
				{
					UINT uIndex = UI_COLLECTION_INVALIDINDEX;
					decltype(auto) e = m_vParams[m_uItemSelected];
					for (size_t i = 0; i < MaxList.size(); i++)
					{
						float fC = *(MaxList.begin() + i);
						if (std::get<3>(e) == fC)
							uIndex = (UINT)i;
					}
					if (uIndex == UI_COLLECTION_INVALIDINDEX)
					{
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_MAXVALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
						hr = E_FAIL;
					}
					else
					{
						hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, uIndex, pPropvarNewValue);
					}
				}
			}
			else if (key == UI_PKEY_StringValue)
			{
				BSTR bstr;
				decltype(auto) e = m_vParams[m_uItemSelected];
				hr = VarBstrFromR4(std::get<3>(e), GetUserDefaultLCID(), 0, &bstr);
				if (FAILED(hr))
					return hr;
				hr = UIInitPropertyFromString(UI_PKEY_StringValue, bstr, pPropvarNewValue);
				SysFreeString(bstr);
			}
			else if (key == UI_PKEY_Enabled)
				hr = UIInitPropertyFromBoolean(UI_PKEY_Enabled, m_bParamSelected, pPropvarNewValue);
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_MINVALUE, [&, MinList](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_Categories)
				hr = S_FALSE;
			else if (key == UI_PKEY_ItemsSource)
			{
				CComPtr<IUICollection> pCollection;
				hr = pPropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr))
					return hr;
				hr = pCollection->Clear();
				for (size_t i = 0; i < MinList.size(); i++)
				{
					CComPtr<CPropertySet> pItem;
					hr = CPropertySet::CreateInstance(&pItem);
					if (FAILED(hr))
						return hr;
					pItem->InitializeItemProperties(NULL, funConvert(*(MinList.begin() + i)).c_str(), UI_COLLECTION_INVALIDINDEX);
					hr = pCollection->Add(pItem);
					hr = S_OK;
				}
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				if (m_bParamSelected)
				{
					UINT uIndex = UI_COLLECTION_INVALIDINDEX;
					decltype(auto) e = m_vParams[m_uItemSelected];
					for (size_t i = 0; i < MinList.size(); i++)
					{
						float fC = *(MinList.begin() + i);
						if (std::get<2>(e) == fC)
							uIndex = (UINT)i;
					}
					if (uIndex == UI_COLLECTION_INVALIDINDEX)
					{
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_MINVALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
						hr = E_FAIL;
					}
					else
					{
						hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, uIndex, pPropvarNewValue);
					}
				}
			}
			else if (key == UI_PKEY_StringValue)
			{
				BSTR bstr;
				decltype(auto) e = m_vParams[m_uItemSelected];
				hr = VarBstrFromR4(std::get<2>(e), GetUserDefaultLCID(), 0, &bstr);
				if (FAILED(hr))
					return hr;
				hr = UIInitPropertyFromString(UI_PKEY_StringValue, bstr, pPropvarNewValue);
				SysFreeString(bstr);
			}
			else if (key == UI_PKEY_Enabled)
				hr = UIInitPropertyFromBoolean(UI_PKEY_Enabled, m_bParamSelected, pPropvarNewValue);
			return hr;
		});
	pComHandler->AddOnUpdatePropertyListener(IDR_CMD_STEP, [&, StepList](UINT nCmdID, REFPROPERTYKEY key, const PROPVARIANT* pPropvarCurrentValue, PROPVARIANT* pPropvarNewValue) -> HRESULT
		{
			UNREFERENCED_PARAMETER(nCmdID);
			HRESULT hr = E_FAIL;
			if (key == UI_PKEY_Categories)
				hr = S_FALSE;
			else if (key == UI_PKEY_ItemsSource)
			{
				CComPtr<IUICollection> pCollection;
				hr = pPropvarCurrentValue->punkVal->QueryInterface(IID_PPV_ARGS(&pCollection));
				if (FAILED(hr))
					return hr;
				hr = pCollection->Clear();
				for (size_t i = 0; i < StepList.size(); i++)
				{
					CComPtr<CPropertySet> pItem;
					hr = CPropertySet::CreateInstance(&pItem);
					if (FAILED(hr))
						return hr;
					pItem->InitializeItemProperties(NULL, funConvert(*(StepList.begin() + i)).c_str(), UI_COLLECTION_INVALIDINDEX);
					hr = pCollection->Add(pItem);
					hr = S_OK;
				}
			}
			else if (key == UI_PKEY_SelectedItem)
			{
				if (m_bParamSelected)
				{
					UINT uIndex = UI_COLLECTION_INVALIDINDEX;
					decltype(auto) e = m_vParams[m_uItemSelected];
					for (size_t i = 0; i < StepList.size(); i++)
					{
						float fC = *(StepList.begin() + i);
						if (std::get<4>(e) == fC)
							uIndex = (UINT)i;
					}
					if (uIndex == UI_COLLECTION_INVALIDINDEX)
					{
						hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_STEP, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_StringValue);
						hr = E_FAIL;
					}
					else
					{
						hr = UIInitPropertyFromUInt32(UI_PKEY_SelectedItem, uIndex, pPropvarNewValue);
					}
				}
			}
			else if (key == UI_PKEY_StringValue)
			{
				BSTR bstr;
				decltype(auto) e = m_vParams[m_uItemSelected];
				hr = VarBstrFromR4(std::get<4>(e), GetUserDefaultLCID(), 0, &bstr);
				if (FAILED(hr))
					return hr;
				hr = UIInitPropertyFromString(UI_PKEY_StringValue, bstr, pPropvarNewValue);
				SysFreeString(bstr);
			}
			else if (key == UI_PKEY_Enabled)
				hr = UIInitPropertyFromBoolean(UI_PKEY_Enabled, m_bParamSelected, pPropvarNewValue);
			return hr;
		});
	ConnectionChanged();
	ParamSelChanged();
	HRESULT hr;
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_MAXVALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_ItemsSource);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_MINVALUE, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_ItemsSource);
	hr = m_RibboFramework.GetFramework()->InvalidateUICommand(IDR_CMD_STEP, UI_INVALIDATIONS_PROPERTY, &UI_PKEY_ItemsSource);
	return 0;
}

void MainWindow::OnTimer(UINT_PTR nIDEvent)
{
	if (nIDEvent == 0)
	{
		if (m_uTargetState < 0 || m_cSignalsCount < 0 || rand() < 3277)
		{
			OnTargetStateChange();
			return;
		}

		int error, state, param;
		if (m_uTargetState == 0x00)
		{
			if (m_FirmPaths.size() > 0)
			{
				decltype(auto) strFirmPath = m_FirmPaths[m_uFirmSel];
				uint8_t* uFirm = new uint8_t[524288];
				uint32_t uFirmLength;
				size_t npos = strFirmPath.rfind('.');
				auto extNm = strFirmPath.substr(npos == std::string::npos ? strFirmPath.length() : npos + 1);
				std::transform(extNm.begin(), extNm.end(), extNm.begin(), tolower);
				if (extNm.compare(L"hex") == 0)
					uFirmLength = ReadHex(strFirmPath, uFirm);
				else
					uFirmLength = ReadBin(strFirmPath, uFirm);
				if (uFirmLength != 0)
				{
					FirmwareUpload(error, uFirm, uFirmLength);
					if (error == 0)
					{
						PurgeComm(hCom, PURGE_TXCLEAR | PURGE_RXCLEAR);
						std::this_thread::sleep_for(std::chrono::milliseconds(100));
						ReadTargetState(error, state, param);
						if (state > 0)
						{
							AddLogString(L"加载固件成功。");
							m_vParamsToSend.clear();
							OnRefreshAllParam();
						}
					}
				}
				else
				{
					AddLogString(L"读取固件失败。");
				}
				delete[] uFirm;
			}
		}
		else
		{
			if (m_vParamsToSend.size() > 0)
			{
				uint16_t uSent;
				WriteParams(error, m_vParamsToSend, uSent);
				if (error != 0)
					AddLogString(L"发送参数失败");
				else
				{
					std::wostringstream stm;
					stm << L"成功发送" << uSent << L"个参数";
					AddLogString(stm.str());
				}
			}
		}

		if (m_uTargetState > 0x00)
		{
			uint8_t* uBufferRead = new uint8_t[65534];
			uint16_t uLenReadBack;

			TextFetch(error, uBufferRead, 65534, uLenReadBack);
			if (error == 0)
			{
				std::string str((char*)uBufferRead, uLenReadBack);
				std::wstring strn = string2wstring(str);

				CHARRANGE cr;
				cr.cpMin = -1;
				cr.cpMax = -1;
				::SendMessage(m_hRichedit, EM_EXSETSEL, 0, (LPARAM)&cr);
				::SendMessage(m_hRichedit, EM_REPLACESEL, 0, (LPARAM)strn.c_str());
				::SendMessage(m_hRichedit, WM_VSCROLL, SB_BOTTOM, 0L);
			}
			else
			{
				if (error > 100)
				{
					std::wostringstream stm;
					stm << L"读取文本失败,通信错误" << (error - 100);
					AddLogString(stm.str());
				}
				else
				{
					std::wostringstream stm;
					stm << L"读取文本失败,下位机返回错误" << error;
					AddLogString(stm.str());
				}
			}

			SignalFetch(error, uBufferRead, 65534, uLenReadBack);
			if (error == 0)
			{
				if (uLenReadBack > 0 && m_cSignalsCount > 0)
				{
					uint8_t* pData = new uint8_t[uLenReadBack];
					memcpy(pData, uBufferRead, uLenReadBack);
					m_DataStore.push(std::make_tuple(pData, uLenReadBack));

					m_iReceiveFromTarget += (uLenReadBack / 4 / m_cSignalsCount);
					std::wostringstream stm;
					stm << L"从目标机累计接收" << m_iReceiveFromTarget << L"组信号";
					AddLogString(stm.str());
				}
			}
			else
			{
				if (error > 100)
				{
					std::wostringstream stm;
					stm << L"读取信号失败,通信错误" << (error - 100);
					AddLogString(stm.str());
				}
				else
				{
					std::wostringstream stm;
					stm << L"读取信号失败,下位机返回错误" << error;
					AddLogString(stm.str());
				}
			}
			delete[] uBufferRead;
		}
	}
}

LRESULT MainWindow::OnRibbonChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	ReLayout();
	bHandled = TRUE;
	return 0;
}

void MainWindow::OnCommand(UINT uNotifyCode, int nID, HWND wndCtl)
{
	if (uNotifyCode == EN_KILLFOCUS && wndCtl == m_hEdit || nID == VK_RETURN && wndCtl == m_hEdit)
	{
		WCHAR wcText[256];
		int len = (int)::SendMessage(m_hEdit, WM_GETTEXT, (WPARAM)256, (LPARAM)wcText);
		if (len != 0)
		{
			std::get<0>(m_vParams[m_uItemSelected]) = std::wstring(wcText, len);
			funDrawItem(this, m_uItemSelected);
		}
		::DestroyWindow(m_hEdit);
		m_hEdit = nullptr;
	}
}

LRESULT MainWindow::OnDBClick(LPNMHDR pnmh)
{
	if (pnmh->hwndFrom != m_hListView)
		return 0;
	LPNMLISTVIEW pnmv = reinterpret_cast<LPNMLISTVIEW>(pnmh);
	if (pnmv->iItem >= 0)
	{
		CRect rc;
		BOOL b = ListView_GetItemRect(m_hListView, pnmv->iItem, &rc, LVIR_BOUNDS);
		int col = ListView_GetColumnWidth(m_hListView, 0);
		int col1 = ListView_GetColumnWidth(m_hListView, 1);
		if (b == TRUE && m_hEdit == nullptr)
		{
			rc.left += col;
			rc.right = rc.left + col1;
			{
				m_hEdit = CreateWindow(WC_EDIT, L"", WS_BORDER | WS_CHILD | WS_VISIBLE, rc.left, rc.top, rc.Width(), rc.Height(), m_hListView, nullptr, GetModuleHandle(nullptr), nullptr);
				if (m_hEdit != nullptr)
				{
					decltype(auto) hFontOld = GetWindowFont(m_hEdit);
					LOGFONT lfText = {};
					SystemParametersInfoForDpi(SPI_GETICONTITLELOGFONT, sizeof(lfText), &lfText, FALSE, GetDpiForSystem());
					HFONT hFontNew = CreateFontIndirect(&lfText);
					if (hFontNew)
					{
						DeleteObject(hFontOld);
						::SendMessage(m_hEdit, WM_SETFONT, (WPARAM)hFontNew, MAKELPARAM(TRUE, 0));
					}
					::SendMessage(m_hEdit, WM_SETTEXT, (WPARAM)0, (LPARAM)std::get<0>(m_vParams[pnmv->iItem]).c_str());
					hMainwnd = m_hWnd;
					hEdit = m_hEdit;
					lEditProc = ::GetWindowLongPtr(m_hEdit, GWLP_WNDPROC);
					::SetWindowLongPtr(m_hEdit, GWLP_WNDPROC, (LONG_PTR)NewEditWndProc);
					::SendMessage(m_hEdit, EM_SETSEL, (WPARAM)0, (LPARAM)(-1));
					::SetFocus(m_hEdit);
				}
			}
		}
	}
	return 0;
}

LRESULT MainWindow::OnSelectChange(LPNMHDR pnmh)
{
	if (pnmh->hwndFrom != m_hListView)
		return 0;
	LPNMLISTVIEW pnmv = reinterpret_cast<LPNMLISTVIEW>(pnmh);
	if ((pnmv->uNewState ^ pnmv->uOldState) & LVIS_SELECTED)
	{
		auto sel = ListView_GetNextItem(m_hListView, -1, LVNI_SELECTED);
		if (sel == -1)
			m_bParamSelected = FALSE;
		else
		{
			m_bParamSelected = TRUE;
			m_uItemSelected = sel;
		}
		ParamSelChanged();
	}
	return 0;
}