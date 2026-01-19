#include "QExplorerFinder.h"
#include "SettingsDialog.h"
#include <QtWidgets/QApplication>
#include <QtNetwork/QLocalServer>
#include <QtNetwork/QLocalSocket>
#include <QtWidgets/QSystemTrayIcon>
#include <QtWidgets/QMenu>
#include <QtGui/QIcon>
#include <QAbstractNativeEventFilter>
#include <QKeySequence>
#include <QSettings>
#include <QTranslator>
#include <QLocale>
#include <QLibraryInfo>

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <exdisp.h>
#include <servprov.h>
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
    window->setWindowTitle(QObject::tr("Finding in: ") + path);
    
    // Always show the window first
    window->show();
    
    // Force window to foreground mechanism
    HWND hwnd = (HWND)window->winId();
    DWORD foregroundThread = GetWindowThreadProcessId(GetForegroundWindow(), NULL);
    DWORD appThread = GetCurrentThreadId();

    if (foregroundThread != appThread) {
        AttachThreadInput(foregroundThread, appThread, TRUE);
        SetForegroundWindow(hwnd);
        SetFocus(hwnd);
        AttachThreadInput(foregroundThread, appThread, FALSE);
    }
    
    window->raise();
    window->activateWindow();
}

QString GetActiveExplorerPath() {
    HWND hwnd = GetForegroundWindow();
    if (!hwnd) return QString();

    wchar_t className[256];
    if (!GetClassNameW(hwnd, className, 256)) return QString();

    if (wcscmp(className, L"CabinetWClass") != 0 && wcscmp(className, L"ExploreWClass") != 0) {
        return QString();
    }
    
    // Get focus info for the Explorer thread
    DWORD dwProcessId;
    DWORD dwThreadId = GetWindowThreadProcessId(hwnd, &dwProcessId);
    GUITHREADINFO guiInfo = { 0 };
    guiInfo.cbSize = sizeof(GUITHREADINFO);
    GetGUIThreadInfo(dwThreadId, &guiInfo);
    HWND hwndFocus = guiInfo.hwndFocus;

    QString foundPath;
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
                            ComPtr<IDispatch> spDispDoc;
                            if (SUCCEEDED(spBrowser->get_Document(&spDispDoc))) {
                                ComPtr<IShellFolderViewDual> spView;
                                if (SUCCEEDED(spDispDoc.As(&spView))) {
                                    BOOL visible = FALSE;
                                    BOOL focused = FALSE;
                                    
                                    ComPtr<IServiceProvider> spSP;
                                    if (SUCCEEDED(spBrowser.As(&spSP))) {
                                        ComPtr<IShellBrowser> spShellBrowser;
                                        if (SUCCEEDED(spSP->QueryService(SID_STopLevelBrowser, IID_IShellBrowser, &spShellBrowser))) {
                                            ComPtr<IShellView> spShellView;
                                            if (SUCCEEDED(spShellBrowser->QueryActiveShellView(&spShellView))) {
                                                HWND hwndView = NULL;
                                                if (SUCCEEDED(spShellView->GetWindow(&hwndView))) {
                                                    if (IsWindowVisible(hwndView)) {
                                                        visible = TRUE;
                                                    }
                                                    if (hwndFocus && (hwndFocus == hwndView || IsChild(hwndView, hwndFocus))) {
                                                        focused = TRUE;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (focused || (visible && foundPath.isEmpty())) {
                                        ComPtr<Folder> spFolder;
                                        if (SUCCEEDED(spView->get_Folder(&spFolder))) {
                                            ComPtr<Folder2> spFolder2;
                                            if (SUCCEEDED(spFolder.As(&spFolder2))) {
                                                ComPtr<FolderItem> spFolderSelf;
                                                if (SUCCEEDED(spFolder2->get_Self(&spFolderSelf))) {
                                                    BSTR bstrPath;
                                                    if (SUCCEEDED(spFolderSelf->get_Path(&bstrPath))) {
                                                        QString tempPath = QString::fromWCharArray(bstrPath);
                                                        SysFreeString(bstrPath);
                                                        
                                                        // If focused, this is definitely the one.
                                                        if (focused) {
                                                            foundPath = tempPath;
                                                            // We can break the outer loop if we found the focused one.
                                                            // But break logic is tricky inside nested blocks.
                                                            // We'll set foundPath and break later.
                                                        } else {
                                                            // If not focused, but visible, keep it as candidate 
                                                            // ONLY if we haven't found a focused one yet.
                                                            // And if we haven't found any visible one yet (foundPath is empty).
                                                            foundPath = tempPath;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (focused) break;
                                    }
                                }
                            }
                            if (foundPath.isEmpty()) {
                                // Keep searching
                            } else {
                                // If focused, stop. If just visible, continue to check if another is focused?
                                // Actually, if we found a focused one, we break immediately.
                                // If we found a visible one, we store it.
                                // We continue loop to see if there is a FOCUSED one.
                                // Wait, the break condition above "if (focused) break;" handles it.
                                // What about the break condition at end of loop?
                            }
                        }
                    }
                }
                VariantClear(&vIndex);
                // If we found the focused one, break entirely.
                // We need a way to check if 'foundPath' came from a focused match.
                // Re-check focused state?
                // Or just: if we found a path, check if we should stop.
                // If we assume only one tab is visible, we could stop. 
                // But the problem is IsWindowVisible might be unreliable.
                // So: Prioritize Focused. If no Focus, fallback to First Visible.
                
                // My logic above:
                // if (focused) break; 
                // But this break only exits the `if ((HWND)hwndBrowser == hwnd)` block? No, it's inside `if` but loop is outside.
                // Wait, `break` exits the nearest loop (the `for` loop). Correct.
                
                // So if focused, we break and use it.
                // If visible, we set foundPath (if empty) and continue.
                // If we later find a focused one, we overwrite foundPath and break.
            }
        }
    }
    CoUninitialize();
    return foundPath;
}

class ExplorerHotkeyFilter : public QAbstractNativeEventFilter {
public:
    ExplorerHotkeyFilter() {
        updateHotkey();
    }
    ~ExplorerHotkeyFilter() {
        UnregisterHotKey(NULL, 8888);
    }

    void updateHotkey() {
        UnregisterHotKey(NULL, 8888);
        QKeySequence seq = SettingsDialog::getHotkey();
        if (seq.isEmpty()) return;

        UINT modifiers = 0;
        UINT vk = 0;

        // Simple mapping for QKeySequence to Windows Hotkey
        // Note: For a robust implementation, use a more complete mapper.
        if (seq[0].toCombined() & Qt::ControlModifier) modifiers |= MOD_CONTROL;
        if (seq[0].toCombined() & Qt::AltModifier) modifiers |= MOD_ALT;
        if (seq[0].toCombined() & Qt::ShiftModifier) modifiers |= MOD_SHIFT;
        if (seq[0].toCombined() & Qt::MetaModifier) modifiers |= MOD_WIN;

        // Extract key. Qt key codes for letters are same as ASCII/VK
        int qtKey = seq[0].toCombined() & 0x01FFFFFF;
        if (qtKey >= Qt::Key_F1 && qtKey <= Qt::Key_F12) {
            vk = VK_F1 + (qtKey - Qt::Key_F1);
        } else if (qtKey >= 'A' && qtKey <= 'Z') {
            vk = qtKey;
        } else if (qtKey >= '0' && qtKey <= '9') {
            vk = qtKey;
        } else {
            // Fallback for some common keys
            switch(qtKey) {
                case Qt::Key_Space: vk = VK_SPACE; break;
                case Qt::Key_Return: vk = VK_RETURN; break;
                case Qt::Key_Escape: vk = VK_ESCAPE; break;
                default: vk = qtKey; break;
            }
        }

        RegisterHotKey(NULL, 8888, modifiers, vk);
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

    QSettings settings("ExplorerSelector", "ExplorerSelector");
    QString lang = settings.value("Language", "").toString();

    QTranslator translator;
    bool loaded = false;
    // Define the path to external translations
    QString translationsPath = QCoreApplication::applicationDirPath() + "/translations";

    if (lang.isEmpty()) {
        loaded = translator.load(QLocale::system(), "explorerselector", "_", translationsPath);
    } else {
        loaded = translator.load("explorerselector_" + lang, translationsPath);
    }

    if (loaded) {
        app.installTranslator(&translator);
    }

    // Load Qt standard translations (for dialogs like QColorDialog, QMessageBox, etc.)
    QTranslator qtTranslator;
    bool qtLoaded = false;
    
    // Qt translations usually follow "qt_<lang>.qm" pattern
    if (lang.isEmpty()) {
        // Try system locale
        qtLoaded = qtTranslator.load(QLocale::system(), "qt", "_", translationsPath);
    } else {
        qtLoaded = qtTranslator.load("qt_" + lang, translationsPath);
    }
    
    if (qtLoaded) {
        app.installTranslator(&qtTranslator);
    }

    QString serverName = "ExplorerSelectorServer";
    QLocalSocket socket;
    socket.connectToServer(serverName);
    if (socket.waitForConnected(500)) {
        QStringList args = QCoreApplication::arguments();
        if (args.count() > 1) {
            socket.write(args[1].toUtf8());
            socket.waitForBytesWritten(1000);
        }
        return 0;
    }

    QLocalServer server;
    if (!server.listen(serverName)) {
        server.removeServer(serverName);
        server.listen(serverName);
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

    ExplorerHotkeyFilter hotkeyFilter;
    app.installNativeEventFilter(&hotkeyFilter);

    QSystemTrayIcon trayIcon(QIcon(":/QExplorerFinder/app_icon.png"), &app);
    trayIcon.setToolTip(QObject::tr("Explorer Selector"));
    
    QMenu trayMenu;
    QAction* settingsAction = trayMenu.addAction(QObject::tr("Settings..."));
    
    auto openSettings = [&hotkeyFilter]() {
        SettingsDialog dlg;
        if (dlg.exec() == QDialog::Accepted) {
            hotkeyFilter.updateHotkey();
        }
    };

    QObject::connect(settingsAction, &QAction::triggered, openSettings);
    QObject::connect(&trayIcon, &QSystemTrayIcon::activated, [&](QSystemTrayIcon::ActivationReason reason) {
        if (reason == QSystemTrayIcon::DoubleClick) {
            openSettings();
        }
    });

    trayMenu.addSeparator();
    QAction* exitAction = trayMenu.addAction(QObject::tr("Exit"));
    QObject::connect(exitAction, &QAction::triggered, &app, &QApplication::quit);
    
    trayIcon.setContextMenu(&trayMenu);
    trayIcon.show();

    QStringList args = QCoreApplication::arguments();
    if (args.count() > 1) {
        openWindow(args[1]);
    }

    return app.exec();
}