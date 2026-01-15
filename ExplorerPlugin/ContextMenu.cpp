#include "pch.h"
#include "ContextMenu.h"
#include <shellapi.h>

extern HINSTANCE g_hInst;
extern long g_cDllRef;

CContextMenu::CContextMenu() : m_cRef(1)
{
    InterlockedIncrement(&g_cDllRef);
}

CContextMenu::~CContextMenu()
{
    InterlockedDecrement(&g_cDllRef);
}

// IUnknown
IFACEMETHODIMP CContextMenu::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] =
    {
        QITABENT(CContextMenu, IShellExtInit),
        QITABENT(CContextMenu, IContextMenu),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) CContextMenu::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) CContextMenu::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }
    return cRef;
}

// IShellExtInit
IFACEMETHODIMP CContextMenu::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    // If pidlFolder is provided, it's likely a background click in a folder
    if (pidlFolder)
    {
        wchar_t szPath[MAX_PATH];
        if (SHGetPathFromIDListW(pidlFolder, szPath))
        {
            m_szSelectedFolder = szPath;
            return S_OK;
        }
    }

    // If pDataObj is provided, it might be a click on a folder item
    if (pDataObj)
    {
        FORMATETC fmt = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM stg = { TYMED_HGLOBAL };
        if (SUCCEEDED(pDataObj->GetData(&fmt, &stg)))
        {
            HDROP hDrop = (HDROP)GlobalLock(stg.hGlobal);
            if (hDrop)
            {
                UINT uNumFiles = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
                if (uNumFiles > 0)
                {
                    wchar_t szPath[MAX_PATH];
                    if (DragQueryFileW(hDrop, 0, szPath, MAX_PATH))
                    {
                        m_szSelectedFolder = szPath;
                    }
                }
                GlobalUnlock(stg.hGlobal);
            }
            ReleaseStgMedium(&stg);
            if (!m_szSelectedFolder.empty())
                return S_OK;
        }
    }

    return E_FAIL;
}

// IContextMenu
IFACEMETHODIMP CContextMenu::QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    InsertMenuW(hMenu, indexMenu, MF_STRING | MF_BYPOSITION, idCmdFirst, L"尋找檔案");
    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 1);
}

IFACEMETHODIMP CContextMenu::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    BOOL fEx = FALSE;
    if (pici->cbSize >= sizeof(CMINVOKECOMMANDINFOEX))
    {
        fEx = TRUE;
    }

    if (HIWORD(pici->lpVerb) != 0)
    {
        return E_INVALIDARG;
    }

    if (LOWORD(pici->lpVerb) == 0)
    {
        // Get the DLL path to find the EXE
        wchar_t szDllPath[MAX_PATH];
        GetModuleFileNameW(g_hInst, szDllPath, MAX_PATH);
        PathRemoveFileSpecW(szDllPath);
        
        std::wstring exePath = std::wstring(szDllPath) + L"\\ExplorerFinder.exe";

        // Construct command line: "PathToExe" "FolderPath"
        std::wstring commandLine = L"\"" + exePath + L"\" \"" + m_szSelectedFolder + L"\"";

        STARTUPINFOW si = { sizeof(si) };
        PROCESS_INFORMATION pi = { 0 };

        // We need a mutable string for CreateProcess
        std::vector<wchar_t> cmdVec(commandLine.begin(), commandLine.end());
        cmdVec.push_back(0);

        if (CreateProcessW(NULL, cmdVec.data(), NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi))
        {
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);
        }
        else
        {
            MessageBoxW(pici->hwnd, L"Failed to start ExplorerFinder.", L"Error", MB_OK | MB_ICONERROR);
        }
    }

    return S_OK;
}

IFACEMETHODIMP CContextMenu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, LPSTR pszName, UINT cchMax)
{
    return E_NOTIMPL;
}
