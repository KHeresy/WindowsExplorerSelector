#include "SettingsDialog.h"
#include <QSettings>
#include <QCoreApplication>
#include <QDir>
#include <QKeySequence>
#include <QMessageBox>

#ifdef Q_OS_WIN
#include <windows.h>
#endif

SettingsDialog::SettingsDialog(QWidget *parent)
    : QDialog(parent)
    , ui(new Ui::SettingsDialog())
{
    ui->setupUi(this);

    QSettings settings("ExplorerFinder", "ExplorerFinder");
    
    // Load Hotkey
    QString hotkeyStr = settings.value("GlobalHotkey", "Ctrl+F3").toString();
    ui->keySequenceEdit->setKeySequence(QKeySequence::fromString(hotkeyStr));

    // Load Language
    ui->cmbLanguage->addItem(tr("System Default"), "");
    ui->cmbLanguage->addItem("English", "en_US");
    ui->cmbLanguage->addItem(tr("Traditional Chinese"), "zh_TW");

    QString currentLang = settings.value("Language", "").toString();
    int index = ui->cmbLanguage->findData(currentLang);
    if (index != -1) {
        ui->cmbLanguage->setCurrentIndex(index);
    }

    // Load AutoStart
    bool autoStart = settings.value("AutoStart", false).toBool();
    ui->chkAutoStart->setChecked(autoStart);

    // Load History Size
    int historySize = settings.value("HistorySize", 10).toInt();
    ui->sbHistorySize->setValue(historySize);
}

SettingsDialog::~SettingsDialog()
{
    delete ui;
}

void SettingsDialog::accept()
{
    QSettings settings("ExplorerFinder", "ExplorerFinder");
    
    QString newHotkey = ui->keySequenceEdit->keySequence().toString();
    bool newAutoStart = ui->chkAutoStart->isChecked();
    
    QString oldLang = settings.value("Language", "").toString();
    QString newLang = ui->cmbLanguage->currentData().toString();
    
    int newHistorySize = ui->sbHistorySize->value();

    settings.setValue("GlobalHotkey", newHotkey);
    settings.setValue("AutoStart", newAutoStart);
    settings.setValue("Language", newLang);
    settings.setValue("HistorySize", newHistorySize);

    if (oldLang != newLang) {
        QMessageBox::information(this, tr("Language Changed"), tr("Please restart the application for the language change to take effect."));
    }

    applyAutoStart(newAutoStart);

    QDialog::accept();
}

void SettingsDialog::on_btnClearHistory_clicked()
{
    QSettings settings("ExplorerFinder", "ExplorerFinder");
    settings.remove("SearchHistory");
    QMessageBox::information(this, tr("History Cleared"), tr("Search history has been cleared."));
}

void SettingsDialog::applyAutoStart(bool enable)
{
#ifdef Q_OS_WIN
    QSettings bootSettings("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", QSettings::NativeFormat);
    QString appPath = QDir::toNativeSeparators(QCoreApplication::applicationFilePath());
    if (enable) {
        bootSettings.setValue("ExplorerFinder", "\"" + appPath + "\"");
    } else {
        bootSettings.remove("ExplorerFinder");
    }
#endif
}

QKeySequence SettingsDialog::getHotkey()
{
    QSettings settings("ExplorerFinder", "ExplorerFinder");
    return QKeySequence::fromString(settings.value("GlobalHotkey", "Ctrl+F3").toString());
}
