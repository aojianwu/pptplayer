#include "pptvbs.h"

#include <QProcess>
#include <QStringList>

#include <qt_windows.h>

#include <QTimer>
#include <QThread>

#include <QAxObject>
#include <QAxWidget>

#include <QDebug>
#include <QFileDialog>

#include "pptplayer.h"

#include "pptcache.h"


pptvbs::pptvbs(QWidget *parent)
	: QDialog(parent)
{
	ui.setupUi(this);
	PPTCache::Get()->init();
}

void pptvbs::on_btn_init_clicked()
{
	PPTCache::Get()->init(ui.comboBox->currentIndex());
}

void pptvbs::on_btn_start_clicked() {
	QString fileName = QFileDialog::getOpenFileName(this,
		tr("Open Image"), QDir::homePath(), tr("PPT Files (*.ppt *.pptx)"));
	player = new PPTPlayer(fileName);
	player->procStart(ui.widget, ui.widget_3);
}


void pptvbs::on_btn_gotoslide_clicked()
{
	int index = ui.lineEdit->text().toInt();
	player->procGotoSlide(index);
}


void pptvbs::on_btn_close_clicked()
{
	player->procClose();
}

void pptvbs::on_btn_prev_clicked()
{
	player->procPrev();
}

void pptvbs::on_btn_next_clicked()
{
	player->procNext();
}
