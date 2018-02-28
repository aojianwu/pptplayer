#ifndef PPTCACHE_H
#define PPTCACHE_H

#include <qt_windows.h>

#include <QAxObject>
#include <vector>

#include <QList>

class PPTCache
{
public:


	bool init(int type = 0);

	QAxObject* getPPTApp();

	HWND getLastPPTWnd();

public:
	static PPTCache* Get();

private:
	QAxObject* m_controller;	// Application
	int  m_runType;

	QList<HWND> m_allPPTWnds;

private:
	PPTCache();
	~PPTCache();

private:
	static PPTCache* inst;
};

#endif // PPTCACHE_H
