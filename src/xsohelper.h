#pragma once

#include <QString>

class XsoHelper
{
public:
	XsoHelper();
	~XsoHelper();

	static QString getCLSIDFromFileName(QString fileName);

	static QString getDefaultAppFromFileName(QString fileName);

	static bool checkInstalled(QString clsid);

	static QString getAPPFromCLSID(QString clsid);

	static bool isWPS(QString appName);


	static const QString APP_MS_PPT;
	static const QString APP_WPS_PPT;
};
