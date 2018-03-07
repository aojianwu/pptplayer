#include <QDebug>
#include <QDir>
#include "pptplayer.h"

#include "pptcache.h"

#include "xsohelper.h"

PPTPlayer::PPTPlayer(QString filepath)
	: opened(NULL),
	m_controller(NULL),
	m_filepath(filepath) {

	init();

}


PPTPlayer::~PPTPlayer() {
	if (opened)
		opened->querySubObject("Close()");
}



bool PPTPlayer::init() {


	if (m_controller != nullptr)
	{
		delete m_controller;
	}

	m_controller = new QAxObject();

	QString appName = XsoHelper::getDefaultAppFromFileName(m_filepath);

	m_controller->setControl(appName);

	m_controller->setProperty("Visible", false);

	return true;
}


bool PPTPlayer::procRun(QWidget* widget, QWidget* fitWidget) {
	presentation = m_controller->querySubObject("Presentations");

	if (!presentation)
	{
		return false;
	}

	opened = presentation->querySubObject("Open(QString, QVariant, QVariant, QVariant)", m_filepath, 1, 1, 0);
	if (!opened) {
		return false;
	}
	opened->setProperty("IsFullScreen", false);

	sss = opened->querySubObject("SlideShowSettings");
	if (!sss) {
		return false;
	}
	sss->setProperty("ShowWithAnimation", true);
	sss->setProperty("LoopUntilStopped", true);
	

	sss->querySubObject("Run()");
	window = opened->querySubObject("SlideShowWindow");
	if (!window) {
		return false;
	}


	auto view = window->querySubObject("View");
	if (!view) {
		return false;
	}

	view->setProperty("ZoomToFit", true);

	// 以磅为单位
	int sw = window->property("Width").toInt();
	int sh = window->property("Height").toInt();

	int w = widget->width();
	int h = widget->height();

	int dw = 0;
	int dh = 0;
	if (sw/sh < w/h)
	{
		dh = h* 0.75;;
		dw = sw * dh / sh;
	}
	else {
		dw = w * 0.75;
		dh = sh * dw / sw;
	}

	window->setProperty("Width", dw);
	window->setProperty("Height", dh);
	window->setProperty("Top", 0);
	window->setProperty("Left", 0);

	// WPS设置该项无效
	//window->setProperty("Top", dtop);
	//window->setProperty("Left", 0);

	HWND pptWnd = PPTCache::Get()->getLastPPTWnd();
	//::SetParent(pptWnd, (HWND)widget->winId());

	int nw = window->property("Width").toInt() * 4 / 3;
	int nh = window->property("Height").toInt() * 4 / 3;

	int top = (h - nh ) / 2;
	int left = (w - nw) / 2;


	fitWidget->move(left, top);
	fitWidget->setFixedSize(nw, nh);


	::SetParent(pptWnd, (HWND)fitWidget->winId());
	return true;
}


bool PPTPlayer::procExport(int index, QString savePath)
{
	auto slides = window->querySubObject("Presentation");
	slides = slides->querySubObject("Slides(int)", index);
	if (!slides) {
		return false;
	}

	qDebug() << savePath;
	slides->querySubObject("Export(QString, QString, int, int)", savePath, "jpg", 1920, 1080);
	if (!slides) {
		return false;
	}
	return true;
}

bool PPTPlayer::procNext() {
	auto view = window->querySubObject("View");
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
	auto view = window->querySubObject("View");
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
	opened->dynamicCall("Close()");
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
	window->dynamicCall("Activate()");

	return true;
}

bool PPTPlayer::procGotoSlide(int index) {


	if (index > getTotalSlides())
	{
		return false;
	}
	auto view = window->querySubObject("View");
	if (!view)
		return false;
	view->dynamicCall("GotoSlide(int, bool)", index, false);

	return true;
}


int PPTPlayer::getCurrentSlide() {
	auto view = window->querySubObject("View");
	if (!view)
		return -1;

	auto slide = view->querySubObject("Slide");
	if (!slide)
		return -1;
	int currentIndex = slide->property("SlideIndex").toInt();

	return currentIndex;
}

int PPTPlayer::getTotalSlides() {
	
	auto slides = opened->querySubObject("Slides");

	if (!slides)
		return -1;
	int count = slides->property("Count").toInt();

	return count;
}