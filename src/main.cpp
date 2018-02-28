#include "pptvbs.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
	QApplication a(argc, argv);
	pptvbs w;
	w.show();
	return a.exec();
}
