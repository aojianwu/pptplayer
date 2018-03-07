#include "xsohelper.h"

#include <QFileInfo>
#include <QSettings>

#include <qt_windows.h>
#include <QDebug>

const QString XsoHelper::APP_MS_PPT = "Powerpoint.Application";
const QString XsoHelper::APP_WPS_PPT = "KWPP.Application";

XsoHelper::XsoHelper()
{
}

XsoHelper::~XsoHelper()
{
}

QString XsoHelper::getCLSIDFromFileName(QString fileName)
{
	HRESULT			hr;
	CLSID           clsid;
	CLSID           clsidConv;
	QString strClsid;

	if (FAILED(GetClassFile((LPWSTR)fileName.toStdWString().c_str(), &clsid)))
	{
		return "";
	}

	if (SUCCEEDED(OleGetAutoConvert(clsid, &clsidConv)))
		clsid = clsidConv;

	wchar_t* string;
	// Get String from CLSID
	if (SUCCEEDED(::StringFromCLSID(clsid, &string)))
	{
		strClsid = QString::fromWCharArray(string);
		::CoTaskMemFree(string);
	}

	return strClsid;
}


bool XsoHelper::checkInstalled(QString clsid)
{
	QString openKey = QString("HKEY_CLASSES_ROOT\\%1\\shell\\open\\command").arg(clsid);
	QSettings openSettings(openKey, QSettings::NativeFormat);
	QString cmds = openSettings.value(".").toString();

	QStringList args = cmds.split("\" \"");
	if (args.size() > 0)
	{
		QString exepath = args.at(0);
		exepath = exepath.replace("\"", "");
		QFile file;
		if (!file.exists(exepath))
		{
			return false;
		}
	}

	return true;
}

QString XsoHelper::getAPPFromCLSID(QString clsid)
{
	QHash<QString, QString> officeMap;
	// MS OFFICE
	officeMap["PowerPoint.Show"] = XsoHelper::APP_MS_PPT;

	// WPS OFFICE
	officeMap["WPP.PPT"] = XsoHelper::APP_WPS_PPT;
	officeMap["KWPP.Presentation"] = XsoHelper::APP_WPS_PPT;

	QHashIterator<QString, QString> it(officeMap);
	while (it.hasNext()) {
		it.next();
		if (clsid.startsWith(it.key()))
		{
			return it.value();
		}
	}

	return "";
}

QString XsoHelper::getDefaultAppFromFileName(QString fileName)
{
	QFileInfo fi(fileName);

	QString suffix = fi.completeSuffix();

	QString filetypeKey = QString("HKEY_CLASSES_ROOT\\.%1").arg(suffix);
	QSettings filetypeSettings(filetypeKey,	QSettings::NativeFormat);
	
	QString defaultApp = filetypeSettings.value(".").toString();
	if (checkInstalled(defaultApp)) {
		return getAPPFromCLSID(defaultApp);
	}

	QStringList subGroups = filetypeSettings.childGroups();

	for each (QString subgroup in subGroups)
	{
		if (checkInstalled(subgroup)) {
			return getAPPFromCLSID(subgroup);
		}
	}

	return "";
}

bool XsoHelper::isWPS(QString appName)
{
	return appName.compare(APP_WPS_PPT, Qt::CaseInsensitive);
}
