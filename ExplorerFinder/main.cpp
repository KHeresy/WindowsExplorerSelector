#include "QExplorerFinder.h"
#include <QtWidgets/QApplication>
#include <QtNetwork/QLocalServer>
#include <QtNetwork/QLocalSocket>
#include <QtWidgets/QSystemTrayIcon>
#include <QtWidgets/QMenu>
#include <QtGui/QIcon>
#include <QAbstractNativeEventFilter>

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <exdisp.h>
#include <wrl/client.h>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")

using namespace Microsoft::WRL;

void openWindow(const QString& path) {
    QExplorerFinder* window = new QExplorerFinder();
    window->setAttribute(Qt::WA_DeleteOnClose);
    window->setTargetPath(path);
    window->setWindowTitle("Finding in: " + path);
    window->show();
    window->raise();
    window->activateWindow();
}

QString GetActiveExplorerPath() {
    HWND hwnd = GetForegroundWindow();
    if (!hwnd) return QString();

    wchar_t className[256];
    if (!GetClassNameW(hwnd, className, 256)) return QString();

    // Explorer windows are usually CabinetWClass or ExploreWClass
    if (wcscmp(className, L"CabinetWClass") != 0 && wcscmp(className, L"ExploreWClass") != 0) {
        return QString();
    }

    QString foundPath;
    
    // Iterate ShellWindows to find this HWND
    CoInitialize(NULL); 
    {
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
                        LONG_PTR hwndBrowser = 0;
                        spBrowser->get_HWND(&hwndBrowser);
                        if ((HWND)hwndBrowser == hwnd) {
                            // Match! Get Path
                            ComPtr<IDispatch> spDispDoc;
                            if (SUCCEEDED(spBrowser->get_Document(&spDispDoc))) {
                                ComPtr<IShellFolderViewDual> spView;
                                if (SUCCEEDED(spDispDoc.As(&spView))) {
                                    ComPtr<Folder> spFolder;
                                    if (SUCCEEDED(spView->get_Folder(&spFolder))) {
                                        ComPtr<Folder2> spFolder2;
                                        if (SUCCEEDED(spFolder.As(&spFolder2))) {
                                            ComPtr<FolderItem> spFolderSelf;
                                            if (SUCCEEDED(spFolder2->get_Self(&spFolderSelf))) {
                                                BSTR bstrPath;
                                                if (SUCCEEDED(spFolderSelf->get_Path(&bstrPath))) {
                                                    foundPath = QString::fromWCharArray(bstrPath);
                                                    SysFreeString(bstrPath);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            break; // Found and processed
                        }
                    }
                }
                VariantClear(&vIndex);
            }
        }
    }
    CoUninitialize();
    
    return foundPath;
}

class ExplorerHotkeyFilter : public QAbstractNativeEventFilter {
public:
    ExplorerHotkeyFilter() {
        // Register Ctrl + F3 (ID: 8888)
        RegisterHotKey(NULL, 8888, MOD_CONTROL, VK_F3);
    }
    ~ExplorerHotkeyFilter() {
        UnregisterHotKey(NULL, 8888);
    }

    bool nativeEventFilter(const QByteArray &eventType, void *message, qintptr *result) override {
        Q_UNUSED(result);
        if (eventType == "windows_generic_MSG") {
            MSG* msg = static_cast<MSG*>(message);
            if (msg->message == WM_HOTKEY && msg->wParam == 8888) {
                QString path = GetActiveExplorerPath();
                if (!path.isEmpty()) {
                    openWindow(path);
                }
                return true;
            }
        }
        return false;
    }
};

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    app.setQuitOnLastWindowClosed(false);

    QString serverName = "ExplorerFinderServer";
    
    // Try to connect to existing server
    QLocalSocket socket;
    socket.connectToServer(serverName);
    if (socket.waitForConnected(500)) {
        QStringList args = QCoreApplication::arguments();
        if (args.count() > 1) {
            socket.write(args[1].toUtf8());
            socket.waitForBytesWritten(1000);
        }
        return 0; // Exit secondary instance
    }

    // Start Server
    QLocalServer server;
    if (!server.listen(serverName)) {
        // If listen fails (maybe stale lock file), try removing and listening again
        server.removeServer(serverName);
        if (!server.listen(serverName)) {
            // If still fails, just run as standalone (or handle error)
        }
    }

    QObject::connect(&server, &QLocalServer::newConnection, [&server]() {
        QLocalSocket* clientConnection = server.nextPendingConnection();
        QObject::connect(clientConnection, &QLocalSocket::readyRead, [clientConnection]() {
            QByteArray data = clientConnection->readAll();
            QString path = QString::fromUtf8(data);
            if (!path.isEmpty()) {
                openWindow(path);
            }
        });
        QObject::connect(clientConnection, &QLocalSocket::disconnected, clientConnection, &QLocalSocket::deleteLater);
    });

    // Install Hotkey Filter
    ExplorerHotkeyFilter hotkeyFilter;
    app.installNativeEventFilter(&hotkeyFilter);

    // System Tray
    QSystemTrayIcon trayIcon(QIcon(":/QExplorerFinder/app_icon.png"), &app);
    trayIcon.setToolTip("Explorer Finder");
    
    QMenu trayMenu;
    QAction* exitAction = trayMenu.addAction("Exit");
    QObject::connect(exitAction, &QAction::triggered, &app, &QApplication::quit);
    
    trayIcon.setContextMenu(&trayMenu);
    trayIcon.show();

    // Initial Window if args provided
    QStringList args = QCoreApplication::arguments();
    if (args.count() > 1) {
        openWindow(args[1]);
    }

    return app.exec();
}
