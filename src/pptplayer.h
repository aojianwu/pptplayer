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
	QAxObject* m_presentations;	// Presentations
	QAxObject* m_opened;			// Presentation 
	QAxObject* m_sss;				// SlideShowSettings
	QAxObject* m_window;			// SlideShowWindow
	QString    m_filepath;

	QString    m_tmpPath;

	QAxObject* m_application;	// Application
	int  m_runType;
};

#endif

