#pragma once
#include "SimpleWnd.h"
#include "RibbonFramework.h"

template<class T>
class thread_safe_queue
{
private:
	mutable std::mutex mut;
	std::queue<std::shared_ptr<T>> data_queue;
	std::condition_variable data_con;
public:
	thread_safe_queue() {}
	thread_safe_queue(thread_safe_queue const& other)
	{
		std::lock_guard<std::mutex> lk(other.mut);
		data_queue = other.data_queue;
	}
	void push(T tValue)
	{
		std::shared_ptr<T> data(std::make_shared<T>(std::move(tValue)));
		std::lock_guard<std::mutex> lk(mut);
		data_queue.push(data);
		data_con.notify_one();
	}
	void wait_and_pop(T& tValue)
	{
		std::unique_lock<std::mutex> lk(mut);
		data_con.wait(lk, [this] {return !data_queue.empty(); });
		tValue = std::move(*data_queue.front());
		data_queue.pop();
	}
	std::shared_ptr<T>wait_and_pop()
	{
		std::unique_lock<std::mutex> lk(mut);
		data_con.wait(lk, [this] {return !data_queue.empty(); });
		std::shared_ptr<T> ret(std::make_shared<T>(data_queue.front()));
		data_queue.pop();
		return ret;
	}
	bool try_pop(T& tValue)
	{
		std::lock_guard<std::mutex> lk(mut);
		if (data_queue.empty())
			return false;
		tValue = std::move(*data_queue.front());
		data_queue.pop();
		return true;
	}

	std::shared_ptr<T> try_pop()
	{
		std::lock_guard<std::mutex> lk(mut);
		if (data_queue.empty())
			return std::shared_ptr<T>();
		std::shared_ptr<T> ret(std::make_shared(data_queue.front()));
		data_queue.pop();
		return ret;
	}

	bool empty() const
	{
		std::lock_guard<std::mutex> lk(mut);
		return data_queue.empty();
	}
};

class MainWindow : public CSimpleWnd
{
public:
	MainWindow();
	virtual ~MainWindow();

	void ReLayout();

	void ConnectionChanged();
	void ParamSelChanged();
	void OnTargetStateChange();

	void AddLogString(std::wstring str);
	void OnRefreshAllParam();
	BOOL RefreshComs();

protected:
	int OnCreate(LPCREATESTRUCT lpCreateStruct);
	void OnSize(UINT nType, CSize size);
	void OnDestroy();
	void OnClose();
	void OnPaint(HDC dc);
	BOOL OnInitDialog(HWND wnd, LPARAM lInitParam);
	void OnTimer(UINT_PTR nIDEvent);
	LRESULT OnRibbonChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
	void OnCommand(UINT uNotifyCode, int nID, HWND wndCtl);
	LRESULT OnDBClick(LPNMHDR pnmh);
	LRESULT OnSelectChange(LPNMHDR pnmh);

	BEGIN_MSG_MAP_EX(MainWindow)
		MSG_WM_CREATE(OnCreate)
		MSG_WM_SIZE(OnSize)
		MSG_WM_DESTROY(OnDestroy)
		MSG_WM_CLOSE(OnClose)
		MSG_WM_PAINT(OnPaint)
		MSG_WM_INITDIALOG(OnInitDialog)
		MSG_WM_TIMER(OnTimer)
		MESSAGE_HANDLER(UM_RIBBONVIEWCHANGED, OnRibbonChanged)
		MSG_WM_COMMAND(OnCommand)
		NOTIFY_CODE_HANDLER_EX(NM_DBLCLK, OnDBClick)
		NOTIFY_CODE_HANDLER_EX(LVN_ITEMCHANGED, OnSelectChange)
		REFLECT_NOTIFICATIONS_EX()
	END_MSG_MAP()

public:
	RibbonFramework m_RibboFramework;
	HWND m_hListView = nullptr, m_hListBox = nullptr, m_hRichedit = nullptr, m_hEdit = nullptr;

	std::wstring m_strFileOpen;
	std::vector<std::tuple<std::wstring, float, float, float, float>> m_vParams;//Name Value Min Max Step
	BOOL m_bChangeUnsaved = FALSE;

	std::vector<std::wstring> m_FirmPaths;
	UINT m_uFirmSel = 0;

	std::vector<std::wstring> m_ComPorts;
	std::wstring m_ComSelected;
	BOOL m_bConnected = FALSE;

	BOOL m_bParamSelected = FALSE;
	UINT m_uItemSelected = 0;

	std::vector<std::tuple<uint16_t, float>> m_vParamsToSend;

	uint32_t m_uLogCount = 0;
	USHORT m_hListenPort = 3320;

	INT m_uTargetState = -1;// 0 bootloader, 1 paused, 2 running
	uint8_t m_cSignalsCount = -1;

	SOCKET m_sServer = INVALID_SOCKET;
	WSAEVENT m_hServerEvent;
	SOCKET m_sClient[WSA_MAXIMUM_WAIT_EVENTS] = { INVALID_SOCKET };
	WSAEVENT m_hClientEvent[WSA_MAXIMUM_WAIT_EVENTS];
	std::atomic_char16_t iTotal = 0;

	thread_safe_queue<std::tuple<uint8_t*, uint16_t>> m_DataStore;

	std::thread m_thUpld;
	std::atomic_bool m_bThreadToStop;
	std::atomic_bool m_bThreadRunning;

	std::atomic_int64_t m_iReceiveFromTarget = 0;
	std::atomic_int64_t m_iTransmitToMATLAB = 0;
};