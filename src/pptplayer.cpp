#include <QDebug>
#include <QDir>

#include <QFile>
#include "pptplayer.h"

#include "pptcache.h"

#include "xsohelper.h"

enum MsoTriState
{
	msoCTrue =	1,
	msoFalse =	0,
	msoTrue = -1,
	msoTriStateMixed =  - 2	,
	msoTriStateToggle =  - 3,	
};

PPTPlayer::PPTPlayer(QString filepath)
	: m_opened(NULL),
	m_application(NULL),
	m_filepath(filepath),
	m_tmpPath(filepath){

	init();

}


PPTPlayer::~PPTPlayer() {
	if (m_opened)
		m_opened->querySubObject("Close()");
}



bool PPTPlayer::init() {


	if (m_application != nullptr)
	{
		delete m_application;
	}


	QFileInfo fi(m_filepath);
	m_tmpPath.replace(fi.baseName(), fi.baseName() + "_tmp");
	bool ret = QFile::copy(m_filepath, m_tmpPath);


	QFileInfo fi2(m_tmpPath);
	//fi2.permission(QFile::WriteUser | QFile::WriteOwner | QFile::WriteGroup | QFile::WriteOther);
	SetFileAttributes((LPCWSTR)m_tmpPath.utf16(), FILE_ATTRIBUTE_NORMAL);


	m_application = new QAxObject();

	QString appName = XsoHelper::getDefaultAppFromFileName(m_tmpPath);

	m_application->setControl(appName);

	//m_application->setProperty("Visible", msoFalse);

	return true;
}


bool PPTPlayer::procRun(QWidget* widget, QWidget* fitWidget) {
	m_presentations = m_application->querySubObject("Presentations");

	if (!m_presentations)
	{
		return false;
	}

	//m_opened = m_presentations->querySubObject("Open(QString, QVariant, QVariant, QVariant)", m_tmpPath, msoTrue, msoTrue, msoFalse);
	m_opened = m_presentations->querySubObject("Open(QString)", m_tmpPath);
	if (!m_opened) {
		return false;
	}
	m_opened->setProperty("IsFullScreen", msoFalse);

	m_sss = m_opened->querySubObject("SlideShowSettings");
	if (!m_sss) {
		return false;
	}
	m_sss->setProperty("ShowWithAnimation", msoTrue);
	m_sss->setProperty("LoopUntilStopped", msoTrue);
	

	m_sss->querySubObject("Run()");
	m_window = m_opened->querySubObject("SlideShowWindow");
	if (!m_window) {
		return false;
	}


	auto view = m_window->querySubObject("View");
	if (!view) {
		return false;
	}

	view->setProperty("ZoomToFit", msoTrue);

	// 以磅为单位
	float sw = m_window->property("Width").toInt();
	float sh = m_window->property("Height").toInt();

	float w = widget->width();
	float h = widget->height();

	float dw = 0;
	float dh = 0;
	if (sw/sh < w/h)
	{
		dh = h* 0.75;;
		dw = sw * dh / sh;
	}
	else {
		dw = w * 0.75;
		dh = sh * dw / sw;
	}

	m_window->setProperty("Width", dw);
	m_window->setProperty("Height", dh);
	m_window->setProperty("Top", 0);
	m_window->setProperty("Left", 0);

	// WPS设置该项无效
	//window->setProperty("Top", dtop);
	//window->setProperty("Left", 0);

	HWND pptWnd = PPTCache::Get()->getLastPPTWnd();
	//::SetParent(pptWnd, (HWND)widget->winId());

	float nw = m_window->property("Width").toFloat() * 4 / 3;
	float nh = m_window->property("Height").toFloat() * 4 / 3;

	float top = (h - nh ) / 2;
	float left = (w - nw) / 2;


	fitWidget->move(left, top);
	fitWidget->setFixedSize(nw, nh);


	::SetParent(pptWnd, (HWND)fitWidget->winId());

	m_opened->dynamicCall("Save()");
	return true;
}


bool PPTPlayer::procExport(int index, QString savePath)
{
	auto slides = m_window->querySubObject("Presentation");
	slides = slides->querySubObject("Slides(int)", index);
	if (!slides) {
		return false;
	}

	qDebug() << savePath;
	slides->querySubObject("Export(QString, QString, int, int)", savePath, "jpg", 1024, 768);
	if (!slides) {
		return false;
	}
	return true;
}

bool PPTPlayer::procNext() {
	auto view = m_window->querySubObject("View");
	if (!view) {
		return false;
	}

	if (getCurrentSlide() == getTotalSlides())
	{
		view->dynamicCall("First()");
	}
	else {
		view->dynamicCall("Next()");
	}

	return true;
}

bool PPTPlayer::procPrev() {
	auto view = m_window->querySubObject("View");
	if (!view)
		return false;
	
	if (getCurrentSlide() == 1)
	{
		view->dynamicCall("Last()");
	}
	else {
		view->dynamicCall("Previous()");
	}

	return true;
}

bool PPTPlayer::procClose() {
	qDebug() << "close";
	if (m_opened)
	{
		m_opened->dynamicCall("Close()");
	}

	if (m_application)
	{
		m_application->dynamicCall("Quit()");
	}

	QFile f;
	f.remove(m_tmpPath);
	
	return true;
}

bool PPTPlayer::procStop() {
	return true;
}

bool PPTPlayer::procStart(QWidget* widget, QWidget* fitWidget) {
	return procRun(widget, fitWidget);
}

bool PPTPlayer::procFocus() {

	// MS Office可用，WPS无效
	m_window->dynamicCall("Activate()");

	return true;
}

bool PPTPlayer::procGotoSlide(int index) {


	if (index > getTotalSlides())
	{
		return false;
	}
	auto view = m_window->querySubObject("View");
	if (!view)
		return false;
	view->dynamicCall("GotoSlide(int, bool)", index, msoFalse);

	return true;
}


int PPTPlayer::getCurrentSlide() {
	auto view = m_window->querySubObject("View");
	if (!view)
		return -1;

	auto slide = view->querySubObject("Slide");
	if (!slide)
		return -1;
	int currentIndex = slide->property("SlideIndex").toInt();

	return currentIndex;
}

int PPTPlayer::getTotalSlides() {
	
	auto slides = m_opened->querySubObject("Slides");

	if (!slides)
		return -1;
	int count = slides->property("Count").toInt();

	return count;
}