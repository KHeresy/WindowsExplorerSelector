#include "SettingsDialog.h"
#include <QSettings>
#include <QCoreApplication>
#include <QDir>
#include <QKeySequence>
#include <QMessageBox>
#include <QProcess>

#ifdef Q_OS_WIN
#include <windows.h>
#endif

SettingsDialog::SettingsDialog(QWidget *parent)
    : QDialog(parent)
    , ui(new Ui::SettingsDialog())
{
    ui->setupUi(this);

    QSettings settings("ExplorerSelector", "ExplorerSelector");
    
    // Load Hotkey
    QString hotkeyStr = settings.value("GlobalHotkey", "Ctrl+F3").toString();
    ui->keySequenceEdit->setKeySequence(QKeySequence::fromString(hotkeyStr));

    // Handle Hotkey Disable (GroupBox checkable)
    if (hotkeyStr.isEmpty()) {
        ui->groupBox->setChecked(false);
    } else {
        ui->groupBox->setChecked(true);
    }

    connect(ui->groupBox, &QGroupBox::toggled, this, [this](bool checked) {
        if (checked && ui->keySequenceEdit->keySequence().isEmpty()) {
            // Restore default if re-enabled and empty
            ui->keySequenceEdit->setKeySequence(QKeySequence("Ctrl+F3"));
        }
    });

    // Load Language
    ui->cmbLanguage->addItem(tr("System Default"), "");
    ui->cmbLanguage->addItem("English", "en_US"); // Source language

    QDir translationsDir(QCoreApplication::applicationDirPath());
    if (translationsDir.cd("translations")) {
        QStringList fileNames = translationsDir.entryList(QStringList() << "explorerselector_*.qm", QDir::Files);
        for (const QString &fileName : fileNames) {
            // Extract locale from filename: explorerselector_zh_TW.qm -> zh_TW
            QString localeStr = fileName.mid(17); // Length of "explorerselector_"
            localeStr.chop(3); // Remove ".qm"

            if (localeStr.compare("en_US", Qt::CaseInsensitive) == 0) continue; // Already added

            QLocale locale(localeStr);
            QString label = locale.nativeLanguageName();
            if (label.isEmpty()) label = localeStr;
            
            // Capitalize first letter
            if (!label.isEmpty()) label[0] = label[0].toUpper();

            ui->cmbLanguage->addItem(label, localeStr);
        }
    }

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

    // Set focus to project link
    ui->lblProjectLink->setFocusPolicy(Qt::StrongFocus);
    ui->lblProjectLink->setTextInteractionFlags(ui->lblProjectLink->textInteractionFlags() | Qt::LinksAccessibleByKeyboard | Qt::TextSelectableByKeyboard);
    QWidget::setTabOrder(ui->lblProjectLink, ui->keySequenceEdit);
    ui->lblProjectLink->setFocus();

    // Connect About Qt button
    connect(ui->btnAboutQt, &QPushButton::clicked, this, [this]() {
        QMessageBox::aboutQt(this);
    });

    // Make window non-resizable
    layout()->setSizeConstraint(QLayout::SetFixedSize);
}

SettingsDialog::~SettingsDialog()
{
    delete ui;
}

void SettingsDialog::accept()
{
    QSettings settings("ExplorerSelector", "ExplorerSelector");
    
    QString newHotkey;
    if (!ui->groupBox->isChecked()) {
        newHotkey = "";
    } else {
        newHotkey = ui->keySequenceEdit->keySequence().toString();
    }

    bool newAutoStart = ui->chkAutoStart->isChecked();
    
    QString oldLang = settings.value("Language", "").toString();
    QString newLang = ui->cmbLanguage->currentData().toString();
    
    int newHistorySize = ui->sbHistorySize->value();

    settings.setValue("GlobalHotkey", newHotkey);
    settings.setValue("AutoStart", newAutoStart);
    settings.setValue("Language", newLang);
    settings.setValue("HistorySize", newHistorySize);

    applyAutoStart(newAutoStart);

    if (oldLang != newLang) {
        auto ret = QMessageBox::question(this, tr("Language Changed"), 
            tr("Language change requires a restart to take effect.\nDo you want to restart now?"),
            QMessageBox::Yes | QMessageBox::No, QMessageBox::Yes);
        
        if (ret == QMessageBox::Yes) {
            QProcess::startDetached(QCoreApplication::applicationFilePath(), QStringList());
            QCoreApplication::quit();
            return;
        }
    }

    QDialog::accept();
}

void SettingsDialog::on_btnClearHistory_clicked()
{
    QSettings settings("ExplorerSelector", "ExplorerSelector");
    settings.remove("SearchHistory");
    QMessageBox::information(this, tr("History Cleared"), tr("Search history has been cleared."));
}

void SettingsDialog::applyAutoStart(bool enable)
{
#ifdef Q_OS_WIN
    QSettings bootSettings("HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Run", QSettings::NativeFormat);
    QString appPath = QDir::toNativeSeparators(QCoreApplication::applicationFilePath());
    if (enable) {
        bootSettings.setValue("ExplorerSelector", "\"" + appPath + "\"");
    } else {
        bootSettings.remove("ExplorerSelector");
    }
#endif
}

QKeySequence SettingsDialog::getHotkey()
{
    QSettings settings("ExplorerSelector", "ExplorerSelector");
    return QKeySequence::fromString(settings.value("GlobalHotkey", "Ctrl+F3").toString());
}