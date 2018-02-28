#include "pptcache.h"

#include <qt_windows.h>

PPTCache::PPTCache() : m_controller(nullptr)
{

}

PPTCache::~PPTCache()
{

}

PPTCache *PPTCache::inst = 0;
PPTCache *PPTCache::Get()
{
	if (!inst)
	{
		inst = new PPTCache;
	}
	return inst;
}

bool PPTCache::init(int type) {
	m_runType = type;
	// 0: "Powerpoint.Application"
	// 1£º"Kwpp.Application"
	QString appName = "Powerpoint.Application";
	if (type == 1) {
		appName = "Kwpp.Application";
	}

	if (m_controller != nullptr)
	{
		delete m_controller;
	}

	m_controller = new QAxObject();

	m_controller->setControl(appName);


	m_controller->setProperty("Visible", false);

	return true;
}

QAxObject* PPTCache::getPPTApp() {
	return m_controller;
}

BOOL CALLBACK __children(HWND hwnd, LPARAM lparam)
{
	std::vector<HWND> * cc = (std::vector<HWND>*)lparam;
	cc->push_back(hwnd);
	return TRUE;
}

std::vector<HWND> children(HWND hwnd)
{
	std::vector<HWND> retval;
	BOOL result = ::EnumChildWindows(hwnd, __children, (LPARAM)&retval);
	return retval;
}

std::string wnd_class_name(HWND hwnd)
{
	char sz[300];
	int len = ::GetClassNameA(hwnd, sz, 299);


	return std::string(sz, len);
}

std::wstring wnd_title(HWND hwnd) {
	wchar_t sz[300];
	int len = ::GetWindowText(hwnd, sz, 299);

	return std::wstring(sz, len);
}

BOOL CALLBACK _topwindow_all_enumerator(HWND hwnd, std::vector<HWND> * pww)
{

	std::string className = wnd_class_name(hwnd);

	if (className == "screenClass")
	{
		pww->push_back(hwnd);
	}
	else if (className == "QWidget")
	{
		std::wstring title = wnd_title(hwnd);
		// WPSÑÝÊ¾ »ÃµÆÆ¬·ÅÓ³
		if (title.find(L"WPSÑÝÊ¾ »ÃµÆÆ¬·ÅÓ³") == 0) {
			pww->push_back(hwnd);
		}
	}

	return true;
}

std::vector<HWND> getAllPPTHwnd()
{
	std::vector<HWND> ww;
	::EnumWindows((WNDENUMPROC)_topwindow_all_enumerator, (LPARAM)&ww);

	return ww;
}

HWND PPTCache::getLastPPTWnd() {
	std::vector<HWND> hwnds = getAllPPTHwnd();

	HWND foundHWND = NULL;
	for each (HWND hwnd in hwnds)
	{
		if (m_allPPTWnds.contains(hwnd)) {
			continue;
		} else {
			foundHWND = hwnd;
		}
	}

	return foundHWND;
}
