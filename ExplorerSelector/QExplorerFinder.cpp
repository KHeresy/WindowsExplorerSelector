#include "QExplorerFinder.h"
#include <QApplication>
#include <QDir>
#include <QStringList>
#include <QDebug>
#include <QSettings>
#include <QKeyEvent>
#include <QTimer>
#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <exdisp.h>
#include <wrl/client.h>

using namespace Microsoft::WRL;

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "user32.lib")

QExplorerFinder::QExplorerFinder(QWidget *parent)
    : QWidget(parent)
    , ui(new Ui::QExplorerFinderClass())
{
    ui->setupUi(this);
    ui->listResults->setSelectionMode(QAbstractItemView::ExtendedSelection);

    setWindowIcon(QIcon(":/QExplorerFinder/app_icon.png"));

    QSettings settings("ExplorerSelector", "ExplorerSelector");
    ui->chkCloseAfterSelect->setChecked(settings.value("CloseAfterSelect", false).toBool());
    ui->chkAlwaysOnTop->setChecked(settings.value("AlwaysOnTop", false).toBool());
    ui->chkDefaultSelectAll->setChecked(settings.value("DefaultSelectAll", false).toBool());

    connect(ui->lineEditSearch->lineEdit(), &QLineEdit::returnPressed, this, &QExplorerFinder::on_lineEditSearch_returnPressed);
    loadHistory();

    m_checkTimer = new QTimer(this);
    connect(m_checkTimer, &QTimer::timeout, this, &QExplorerFinder::onCheckExplorerWindow);
    m_checkTimer->start(1000);
}

QExplorerFinder::~QExplorerFinder()
{
    QSettings settings("ExplorerSelector", "ExplorerSelector");
    settings.setValue("CloseAfterSelect", ui->chkCloseAfterSelect->isChecked());
    settings.setValue("AlwaysOnTop", ui->chkAlwaysOnTop->isChecked());
    settings.setValue("DefaultSelectAll", ui->chkDefaultSelectAll->isChecked());

    delete ui;
}

void QExplorerFinder::setTargetPath(const QString& path)
{
    m_targetPath = path;
    // Normalize path to OS separator
    m_targetPath.replace("/", "\\");
    // Remove trailing backslash if present (unless it's root drive like C:\)
    if (m_targetPath.length() > 3 && m_targetPath.endsWith("\\")) {
        m_targetPath.chop(1);
    }
    ui->lineEditFolder->setText(m_targetPath);
}

void QExplorerFinder::on_btnSearch_clicked()
{
    if (m_targetPath.isEmpty()) return;

    QString searchPattern = ui->lineEditSearch->currentText();
    if (searchPattern.isEmpty()) return;

    saveHistory(searchPattern);

    // Support wildcards if user provided, otherwise *pattern*
    if (!searchPattern.contains("*") && !searchPattern.contains("?")) {
        searchPattern = "*" + searchPattern + "*";
    }
    searchPattern.replace("[", "[[]");

    ui->lblStatus->setText(tr("Searching..."));
    
    ui->listResults->clear();
    QListWidgetItem *item = new QListWidgetItem(tr("Searching..."));
    item->setTextAlignment(Qt::AlignCenter);
    item->setForeground(Qt::gray);
    ui->listResults->addItem(item);
    ui->listResults->setEnabled(false);
    
    QApplication::processEvents();

    QDir dir(m_targetPath);
    QStringList files = dir.entryList(QStringList() << searchPattern, QDir::Files | QDir::Dirs | QDir::NoDotAndDotDot);
    
    ui->listResults->setEnabled(true);
    ui->listResults->clear();
    ui->listResults->addItems(files);
    
    if (ui->chkDefaultSelectAll->isChecked()) {
        ui->listResults->selectAll();
    }
    
    if (ui->listResults->count() > 0) {
        ui->listResults->setFocus();
        if (!ui->chkDefaultSelectAll->isChecked()) {
            ui->listResults->setCurrentRow(0); // Select first item for navigation convenience if not all selected
        }
    }
    
    updateStatusLabel();
}

void QExplorerFinder::on_btnSelectAll_clicked()
{
    QStringList files;
    auto selectedItems = ui->listResults->selectedItems();
    for(auto item : selectedItems)
    {
        files << item->text();
    }
    if (!files.isEmpty()) {
        selectFiles(files);
        if (ui->chkCloseAfterSelect->isChecked()) {
            close();
        }
    }
}

void QExplorerFinder::on_lineEditSearch_returnPressed()
{
    on_btnSearch_clicked();
}

void QExplorerFinder::on_listResults_itemDoubleClicked(QListWidgetItem *item)
{
    if (item) {
        selectFiles(QStringList() << item->text());
        if (ui->chkCloseAfterSelect->isChecked()) {
            close();
        }
    }
}

void QExplorerFinder::on_listResults_itemActivated(QListWidgetItem *item)
{
    Q_UNUSED(item);
    on_btnSelectAll_clicked();
}

void QExplorerFinder::on_listResults_itemSelectionChanged()
{
    updateStatusLabel();
}

void QExplorerFinder::on_btnListSelectAll_clicked()
{
    ui->listResults->selectAll();
}

void QExplorerFinder::on_btnListInvert_clicked()
{
    for(int i = 0; i < ui->listResults->count(); ++i) {
        QListWidgetItem* item = ui->listResults->item(i);
        item->setSelected(!item->isSelected());
    }
}

void QExplorerFinder::on_btnListNone_clicked()
{
    ui->listResults->clearSelection();
}

void QExplorerFinder::on_chkAlwaysOnTop_stateChanged(int state)
{
    Qt::WindowFlags flags = windowFlags();
    if (state == Qt::Checked) {
        flags |= Qt::WindowStaysOnTopHint;
    } else {
        flags &= ~Qt::WindowStaysOnTopHint;
    }
    setWindowFlags(flags);
    show(); // Needed to apply window flags changes
}

void QExplorerFinder::updateStatusLabel()
{
    int total = ui->listResults->count();
    int selected = ui->listResults->selectedItems().count();
    ui->lblStatus->setText(tr("Found: %1 | Selected: %2").arg(total).arg(selected));
}

// Helper to get Folder Object from WebBrowser
HRESULT GetFolderFromShellBrowser(IWebBrowser2* pBrowser, Folder** ppFolder)
{
    if (!pBrowser) return E_POINTER;
    
    ComPtr<IDispatch> spDispDoc;
    HRESULT hr = pBrowser->get_Document(&spDispDoc);
    if (FAILED(hr)) return hr;

    ComPtr<IShellFolderViewDual> spView;
    hr = spDispDoc.As(&spView);
    if (FAILED(hr)) return hr;

    return spView->get_Folder(ppFolder);
}

void QExplorerFinder::selectFiles(const QStringList& files)
{
    HRESULT hr = CoInitialize(NULL);
    
    ComPtr<IShellWindows> spShellWindows;
    hr = CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&spShellWindows));
    if (FAILED(hr)) {
        return;
    }

    long count = 0;
    spShellWindows->get_Count(&count);

    for (long i = 0; i < count; ++i)
    {
        ComPtr<IDispatch> spDisp;
        VARIANT vIndex;
        VariantInit(&vIndex);
        vIndex.vt = VT_I4;
        vIndex.lVal = i;
        
        hr = spShellWindows->Item(vIndex, &spDisp);
        VariantClear(&vIndex);
        
        if (FAILED(hr)) continue;

        ComPtr<IWebBrowser2> spBrowser;
        hr = spDisp.As(&spBrowser);
        if (FAILED(hr)) continue;

        ComPtr<Folder> spFolder;
        if (FAILED(GetFolderFromShellBrowser(spBrowser.Get(), &spFolder))) continue;

        ComPtr<Folder2> spFolder2;
        if (FAILED(spFolder.As(&spFolder2))) continue;

        ComPtr<FolderItem> spFolderSelf;
        if (FAILED(spFolder2->get_Self(&spFolderSelf))) continue;

        BSTR bstrPath;
        if (FAILED(spFolderSelf->get_Path(&bstrPath))) continue;

        QString folderPath = QString::fromWCharArray(bstrPath);
        SysFreeString(bstrPath);

        // Compare paths
        // Normalize folderPath
        folderPath.replace("/", "\\");
        if (folderPath.length() > 3 && folderPath.endsWith("\\")) folderPath.chop(1);

        if (folderPath.compare(m_targetPath, Qt::CaseInsensitive) == 0)
        {
            // Found the window!
            ComPtr<IDispatch> spDispDoc;
            spBrowser->get_Document(&spDispDoc);
            ComPtr<IShellFolderViewDual> spView;
            spDispDoc.As(&spView);

            bool first = true;
            for (const QString& file : files)
            {
                ComPtr<FolderItem> spItem;
                BSTR bstrName = SysAllocString(file.toStdWString().c_str());
                hr = spFolder->ParseName(bstrName, &spItem);
                SysFreeString(bstrName);

                if (SUCCEEDED(hr) && spItem)
                {
                    long flags = 0;
                    if (first) {
                        flags = 1 | 4 | 8 | 16; // Select, Deselect others, Ensure visible, Focus
                        first = false;
                    } else {
                        flags = 1 | 8; // Select, Ensure visible
                    }
                    
                    VARIANT vItem;
                    VariantInit(&vItem);
                    vItem.vt = VT_DISPATCH;
                    vItem.pdispVal = spItem.Get();
                    // spItem.Get() does not AddRef, but VariantInit/Clear doesn't Release if we don't own it?
                    // Wait, VT_DISPATCH in VARIANT usually implies ownership if we use VariantClear?
                    // VariantClear calls Release() if VT_DISPATCH.
                    // So we must AddRef if we put it in Variant.
                    spItem.Get()->AddRef();
                    
                    spView->SelectItem(&vItem, flags);
                    
                    VariantClear(&vItem); // Releases
                }
            }
            
            // Bring window to front
            LONG_PTR hwnd;
            spBrowser->get_HWND(&hwnd);
            SetForegroundWindow((HWND)hwnd);
            break; // Done
        }
    }

    CoUninitialize();
}

void QExplorerFinder::keyPressEvent(QKeyEvent *event)
{
    if (event->key() == Qt::Key_Escape) {
        close();
    } else {
        QWidget::keyPressEvent(event);
    }
}

void QExplorerFinder::loadHistory()
{
    QSettings settings("ExplorerSelector", "ExplorerSelector");
    QStringList history = settings.value("SearchHistory").toStringList();
    ui->lineEditSearch->clear();
    ui->lineEditSearch->addItems(history);
    ui->lineEditSearch->setCurrentIndex(-1);
}

void QExplorerFinder::saveHistory(const QString& text)
{
    QSettings settings("ExplorerSelector", "ExplorerSelector");
    QStringList history = settings.value("SearchHistory").toStringList();
    int maxHistory = settings.value("HistorySize", 10).toInt();

    history.removeAll(text); // Remove duplicates
    history.prepend(text);   // Add to top

    while (history.size() > maxHistory) {
        history.removeLast();
    }

    settings.setValue("SearchHistory", history);
    
    // Update UI without resetting current text if possible, or just re-load
    // Simplest is to block signals and reload, but that might mess with typing?
    // Actually, we usually search *after* typing.
    // Let's just update the list.
    ui->lineEditSearch->blockSignals(true);
    ui->lineEditSearch->clear();
    ui->lineEditSearch->addItems(history);
    ui->lineEditSearch->setEditText(text);
    ui->lineEditSearch->blockSignals(false);
}

void QExplorerFinder::onCheckExplorerWindow()
{
    if (m_targetPath.isEmpty()) return;

    bool found = false;
    HRESULT hr = CoInitialize(NULL);
    if (SUCCEEDED(hr)) {
        ComPtr<IShellWindows> spShellWindows;
        if (SUCCEEDED(CoCreateInstance(CLSID_ShellWindows, NULL, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&spShellWindows)))) {
            long count = 0;
            spShellWindows->get_Count(&count);

            for (long i = 0; i < count; ++i) {
                ComPtr<IDispatch> spDisp;
                VARIANT vIndex;
                VariantInit(&vIndex);
                vIndex.vt = VT_I4;
                vIndex.lVal = i;

                if (SUCCEEDED(spShellWindows->Item(vIndex, &spDisp))) {
                    ComPtr<IWebBrowser2> spBrowser;
                    if (SUCCEEDED(spDisp.As(&spBrowser))) {
                        ComPtr<Folder> spFolder;
                        if (SUCCEEDED(GetFolderFromShellBrowser(spBrowser.Get(), &spFolder))) {
                            ComPtr<Folder2> spFolder2;
                            if (SUCCEEDED(spFolder.As(&spFolder2))) {
                                ComPtr<FolderItem> spFolderSelf;
                                if (SUCCEEDED(spFolder2->get_Self(&spFolderSelf))) {
                                    BSTR bstrPath;
                                    if (SUCCEEDED(spFolderSelf->get_Path(&bstrPath))) {
                                        QString folderPath = QString::fromWCharArray(bstrPath);
                                        SysFreeString(bstrPath);

                                        folderPath.replace("/", "\\");
                                        if (folderPath.length() > 3 && folderPath.endsWith("\\")) folderPath.chop(1);

                                        if (folderPath.compare(m_targetPath, Qt::CaseInsensitive) == 0) {
                                            found = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                VariantClear(&vIndex);
                if (found) break;
            }
        }
        CoUninitialize();
    }

    if (!found) {
        close();
    }
}
