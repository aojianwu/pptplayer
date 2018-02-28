#pragma once

#include <QtWidgets/QDialog>
#include "ui_pptvbs.h"

#include "pptplayer.h"

class pptvbs : public QDialog
{
	Q_OBJECT

public:
	pptvbs(QWidget *parent = Q_NULLPTR);

public slots:
void on_btn_init_clicked();

void on_btn_start_clicked();

void on_btn_gotoslide_clicked();




void on_btn_close_clicked();



void on_btn_prev_clicked();
void on_btn_next_clicked();


private:
	Ui::pptvbsClass ui;

	PPTPlayer *player;
};
