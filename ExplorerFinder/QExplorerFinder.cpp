#include "QExplorerFinder.h"
#include <QDir>
#include <QStringList>
#include <QDebug>
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
}

QExplorerFinder::~QExplorerFinder()
{
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
    ui->editFolder->setText(m_targetPath);
}

void QExplorerFinder::on_pushButton_clicked()
{
    if (m_targetPath.isEmpty()) return;

    QString searchPattern = ui->lineEdit->text();
    if (searchPattern.isEmpty()) return;

    // Support wildcards if user provided, otherwise *pattern*
    if (!searchPattern.contains("*") && !searchPattern.contains("?")) {
        searchPattern = "*" + searchPattern + "*";
    }

    QDir dir(m_targetPath);
    QStringList files = dir.entryList(QStringList() << searchPattern, QDir::Files | QDir::Dirs | QDir::NoDotAndDotDot);
    
    ui->listWidget->clear();
    ui->listWidget->addItems(files);
}

void QExplorerFinder::on_buttonSelected_clicked()
{
    QStringList files;
    for(int i = 0; i < ui->listWidget->count(); ++i)
    {
        files << ui->listWidget->item(i)->text();
    }
    selectFiles(files);
}

void QExplorerFinder::on_lineEdit_returnPressed()
{
    on_pushButton_clicked();
}

void QExplorerFinder::on_listWidget_itemClicked(QListWidgetItem *item)
{
    if (item) {
        selectFiles(QStringList() << item->text());
    }
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