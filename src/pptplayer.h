#ifndef _PLAYER_PPTPLAYER_H_
#define _PLAYER_PPTPLAYER_H_

#include <QAxObject>
#include <QString>

#include <qt_windows.h>
#include <QWidget>



class PPTPlayer  {
public:
	PPTPlayer(QString filepath);

	

	bool procRun(QWidget* widget = nullptr, QWidget* fitWidget = nullptr);
	bool procNext();
	bool procPrev();
	bool procClose();
	bool procStop();
	bool procStart(QWidget* widget = nullptr, QWidget* fitWidget = nullptr);

	bool procExport(int index, QString savePath);

	int getCurrentSlide();
	int getTotalSlides();

	bool procFocus();

	bool procGotoSlide(int index);

	virtual ~PPTPlayer();

private:
	bool init();

protected:
	QAxObject* presentation;	// Presentations
	QAxObject* opened;			// Presentation 
	QAxObject* sss;				// SlideShowSettings
	QAxObject* window;			// SlideShowWindow
	QString    m_filepath;

	QAxObject* m_controller;	// Application
	int  m_runType;
};

#endif

