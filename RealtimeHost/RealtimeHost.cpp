#include "stdafx.h"
#include "SimpleWnd.h"
#include "MainFrm.h"
#define MAX_LOADSTRING 100


int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPTSTR    lpCmdLine,
	_In_ int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);
	HeapSetInformation(nullptr, HeapEnableTerminationOnCorruption, nullptr, 0);
	if (FAILED(CoInitialize(NULL)))
		return -1;
	auto hMsft = LoadLibraryW(L"msftedit.dll");
	if (hMsft == NULL)
		return -1;
//#ifdef DEBUG
//	FILE* fpDebugOut = NULL;
//	FILE* fpDebugIn = NULL;
//	if (!AllocConsole())
//		MessageBox(NULL, _T("控制台生成失败。"), NULL, 0);
//	SetConsoleTitle(_T("Debug Window"));
//	_tfreopen_s(&fpDebugOut, _T("CONOUT$"), _T("w"), stdout);
//	_tfreopen_s(&fpDebugIn, _T("CONIN$"), _T("r"), stdin);
//	_tsetlocale(LC_ALL, _T("chs"));
//#endif // DEBUG

	MSG msg;
	{
		CSimpleWndHelper::Init(hInstance, _T("RTH"));
		MainWindow wnd;
		auto dpi = GetDpiForSystem();
		wnd.Create(_T("实时控制器-主机端"), WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN /*| WS_CLIPSIBLINGS*/, 0, 0, 0, 1000 * dpi / 96, 630 * dpi / 96, 0, 0);
		wnd.SendMessage(WM_INITDIALOG);
		wnd.CenterWindow();
		wnd.ShowWindow(nCmdShow);
		while (GetMessage(&msg, NULL, 0, 0))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
//#ifdef DEBUG
//	fclose(fpDebugOut);
//	fclose(fpDebugIn);
//	FreeConsole();
//#endif // DEBUG
	FreeLibrary(hMsft);
	CoUninitialize();
	return (int)msg.wParam;
}