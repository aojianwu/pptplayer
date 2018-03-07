#ifndef PPTCACHE_H
#define PPTCACHE_H

#include <qt_windows.h>

#include <QAxObject>
#include <vector>

#include <QList>

class PPTCache
{
public:



	HWND getLastPPTWnd();

public:
	static PPTCache* Get();

private:
	QList<HWND> m_allPPTWnds;

private:
	PPTCache();
	~PPTCache();

private:
	static PPTCache* inst;
};

#endif // PPTCACHE_H
