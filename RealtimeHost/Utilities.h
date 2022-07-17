#pragma once


decltype(auto) funBrosweSaveLocation = [=](MainWindow* pMainWnd) {
	CComPtr<IFileDialog> pfd;
	HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pfd));
	COMDLG_FILTERSPEC fileType[] =
	{
		{L"XML files", L"*.xml"},
		{L"All files", L"*.*"} };
	hr = pfd->SetFileTypes(ARRAYSIZE(fileType), fileType);
	hr = pfd->SetFileTypeIndex(1);
	hr = pfd->SetDefaultExtension(L"xml");
	hr = pfd->Show(pMainWnd->m_hWnd);
	CComPtr<IShellItem> pShellItem;
	hr = pfd->GetResult(&pShellItem);
	std::wstring strFilePath;
	if (pShellItem)
	{
		LPWSTR filePath;
		hr = pShellItem->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &filePath);
		strFilePath = filePath;
		CoTaskMemFree(filePath);
		return strFilePath;
	}
	else
		return std::wstring();
};

decltype(auto) funSaveXML = [=](MainWindow* pMainWnd) {
	pugi::xml_document doc;
	decltype(auto) rootNode = doc.append_child(L"Root");
	if (!rootNode)
	{
		MessageBox(pMainWnd->m_hWnd, L"ÎÄµµ±£´æ³ö´í,0x01", L"´íÎó", MB_ICONERROR);
		return;
	}
	for (decltype(auto) index = 0U; index < pMainWnd->m_vParams.size(); index++)
	{
		decltype(auto) ele = pMainWnd->m_vParams[index];
		decltype(auto) cNode = rootNode.append_child(L"Param");
		if (!cNode)
		{
			MessageBox(pMainWnd->m_hWnd, L"ÎÄµµ±£´æ³ö´í,0x02", L"´íÎó", MB_ICONERROR);
			return;
		}
		cNode.append_attribute(L"Name").set_value(std::get<0>(ele).c_str());
		cNode.append_attribute(L"Value").set_value(std::get<1>(ele));
		cNode.append_attribute(L"Min").set_value(std::get<2>(ele));
		cNode.append_attribute(L"Max").set_value(std::get<3>(ele));
		cNode.append_attribute(L"Step").set_value(std::get<4>(ele));
	}
	doc.save_file(pMainWnd->m_strFileOpen.c_str());
	pMainWnd->m_bChangeUnsaved = FALSE;
};

decltype(auto) funGetTime = []() -> std::wstring
{
	WCHAR szCurrentDateTime[32];
	SYSTEMTIME systm;
	GetLocalTime(&systm);
	wsprintf(szCurrentDateTime, L"%.2d:%.2d:%.2d", systm.wHour, systm.wMinute, systm.wSecond);
	return szCurrentDateTime;
};

decltype(auto) funConvert = [](float Num)
{
	std::wostringstream oss;
	oss << Num;
	std::wstring str(oss.str());
	return str;
};

decltype(auto) funConstrain = [](float Num, float min, float max)
{
	if (Num < min)
		return min;
	else if (Num > max)
		return max;
	return Num;
};

decltype(auto) funDrawItem = [](MainWindow* pMwd, int index, BOOL bInsert = FALSE)
{
	WCHAR pszText[256];
	decltype(auto) e = pMwd->m_vParams[index];
	LVITEM vitem;
	vitem.mask = LVIF_TEXT;
	vitem.iItem = index;
	if (bInsert)
	{
		wsprintf(pszText, L"%d", index);
		vitem.pszText = pszText;
		vitem.iSubItem = 0;
		ListView_InsertItem(pMwd->m_hListView, &vitem);
	}
	wcscpy_s(pszText, std::get<0>(e).c_str());
	vitem.pszText = pszText;
	vitem.iSubItem = 1;
	ListView_SetItem(pMwd->m_hListView, &vitem);

	wcscpy_s(pszText, funConvert(std::get<1>(e)).c_str());
	vitem.pszText = pszText;
	vitem.iSubItem = 2;
	ListView_SetItem(pMwd->m_hListView, &vitem);
};

auto GetValue = [](uint8_t uO)->uint8_t
{
	uint8_t urt;
	if ((uO >= '0') && (uO <= '9'))
		urt = uO - 48;
	else if ((uO >= 'a') && (uO <= 'f'))
		urt = uO - 'a' + 10;
	else if ((uO >= 'A') && (uO <= 'F'))
		urt = uO - 'A' + 10;
	else
		urt = 0;
	return urt;
};

auto Convert = [](uint8_t* pBuffer, uint16_t uLength)->uint16_t
{
	if (uLength % 2 != 0)
		return 0;
	uint16_t uc = 0;
	for (uint16_t ui = 0U; ui < (uLength / 2); ui++)
	{
		char c1 = (char)pBuffer[2 * ui];
		char c2 = (char)pBuffer[2 * ui + 1];
		uint8_t u = GetValue(c1) * 16 + GetValue(c2);
		pBuffer[ui] = u;
		uc++;
	}
	return uc;
};

auto ReadHex = [](std::wstring strFirm, uint8_t* puMcuFirm)->uint32_t
{
	uint8_t uBuffer[48];
	HANDLE hHexFile = CreateFile(strFirm.c_str(), GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hHexFile == INVALID_HANDLE_VALUE || hHexFile == nullptr)
		return 0;
	uint32_t uFirmReadPos = 0;
	uint8_t uDatalength = 0;
	do
	{
		DWORD dwReads;
		if (ReadFile(hHexFile, uBuffer, 9, &dwReads, NULL) != TRUE || dwReads != 9 || uBuffer[0] != ':')
		{
			CloseHandle(hHexFile);
			return 0;
		}
		Convert(uBuffer + 1, 8);
		uDatalength = uBuffer[1];
		uint8_t uType = uBuffer[4];
		if (ReadFile(hHexFile, uBuffer, 4 + 2 * uDatalength, &dwReads, NULL) != TRUE || dwReads != 4 + 2 * uDatalength)
		{
			CloseHandle(hHexFile);
			return 0;
		}
		else if (uType == 0)
		{
			Convert(uBuffer, uDatalength * 2);
			memcpy(puMcuFirm + uFirmReadPos, uBuffer, uDatalength);
			uFirmReadPos += uDatalength;
		}

	} while (uDatalength > 0);
	CloseHandle(hHexFile);
	return uFirmReadPos;
};

auto ReadBin = [](std::wstring strFirm, uint8_t* puMcuFirm)->uint32_t
{
	HANDLE hBinFile = CreateFile(strFirm.c_str(), GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hBinFile == INVALID_HANDLE_VALUE || hBinFile == nullptr)
		return 0;
	uint8_t uDatalength = 0;
	uint32_t uBinSize = GetFileSize(hBinFile, nullptr);
	DWORD dwReads;
	if (ReadFile(hBinFile, puMcuFirm, uBinSize, &dwReads, NULL) != TRUE || dwReads != uBinSize)
	{
		CloseHandle(hBinFile);
		return 0;
	}
	CloseHandle(hBinFile);
	return uBinSize;
};