#include "pch.h"
#include "ContextMenu.h"
#include "resource.h"
#include <shellapi.h>
#include <shlwapi.h>
#include <vector>
#include <string>

extern HINSTANCE g_hInst;
extern long g_cDllRef;

HBITMAP IconToBitmapPARGB32(HICON hIcon) {
    if (!hIcon) return NULL;
    ICONINFO ii = {0};
    if (!GetIconInfo(hIcon, &ii)) return NULL;
    
    BITMAP bm;
    GetObject(ii.hbmColor, sizeof(bm), &bm);
    
    BITMAPINFO bi = {0};
    bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth = bm.bmWidth;
    bi.bmiHeader.biHeight = -bm.bmHeight; // Negative for top-down
    bi.bmiHeader.biPlanes = 1;
    bi.bmiHeader.biBitCount = 32;
    bi.bmiHeader.biCompression = BI_RGB;
    
    HDC hDC = GetDC(NULL);
    void* pBits = NULL;
    HBITMAP hBitmap = CreateDIBSection(hDC, &bi, DIB_RGB_COLORS, &pBits, NULL, 0);
    
    if (hBitmap) {
        HDC hMemDC = CreateCompatibleDC(hDC);
        HBITMAP hOldBmp = (HBITMAP)SelectObject(hMemDC, hBitmap);
        
        // Initialize alpha channel to 0
        memset(pBits, 0, bm.bmWidth * bm.bmHeight * 4);
        
        DrawIconEx(hMemDC, 0, 0, hIcon, bm.bmWidth, bm.bmHeight, 0, NULL, DI_NORMAL);
        
        // Fix Alpha using Mask
        HDC hMaskDC = CreateCompatibleDC(hDC);
        HBITMAP hOldMask = (HBITMAP)SelectObject(hMaskDC, ii.hbmMask);
        
        DWORD* pPixels = (DWORD*)pBits;
        for (int y = 0; y < bm.bmHeight; y++) {
            for (int x = 0; x < bm.bmWidth; x++) {
                // Check Mask (0=Opaque, 1=Transparent)
                // GetPixel returns non-zero for white (1), zero for black (0)
                COLORREF maskColor = GetPixel(hMaskDC, x, y);
                if (maskColor == 0) { // Opaque
                    DWORD& pixel = pPixels[y * bm.bmWidth + x];
                    // If alpha is 0, force it to 255 (Opaque)
                    if ((pixel & 0xFF000000) == 0) {
                        pixel |= 0xFF000000;
                    }
                }
            }
        }
        
        SelectObject(hMaskDC, hOldMask);
        DeleteDC(hMaskDC);

        SelectObject(hMemDC, hOldBmp);
        DeleteDC(hMemDC);
    }
    
    ReleaseDC(NULL, hDC);
    DeleteObject(ii.hbmColor);
    DeleteObject(ii.hbmMask);
    return hBitmap;
}

CContextMenu::CContextMenu() : m_cRef(1)
{
    InterlockedIncrement(&g_cDllRef);
}

CContextMenu::~CContextMenu()
{
    if (m_hMenuBitmap) DeleteObject(m_hMenuBitmap);
    InterlockedDecrement(&g_cDllRef);
}

// IUnknown
IFACEMETHODIMP CContextMenu::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] =
    {
        QITABENT(CContextMenu, IShellExtInit),
        QITABENT(CContextMenu, IContextMenu),
        QITABENT(CContextMenu, IExplorerCommand),
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

    // Load Icon
    HICON hIcon = (HICON)LoadImageW(g_hInst, MAKEINTRESOURCEW(IDI_ICON1), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
    HBITMAP hBitmap = NULL;
    if (hIcon) {
        hBitmap = IconToBitmapPARGB32(hIcon);
        DestroyIcon(hIcon);
    }

    // Load Menu Text from resources
    wchar_t szMenuText[128];
    if (0 == LoadStringW(g_hInst, IDS_MENU_TEXT, szMenuText, ARRAYSIZE(szMenuText))) {
        wcscpy_s(szMenuText, L"Find Files"); // Fallback
    }

    MENUITEMINFOW mii = { sizeof(mii) };
    mii.fMask = MIIM_STRING | MIIM_ID | MIIM_BITMAP;
    mii.wID = idCmdFirst;
    mii.dwTypeData = szMenuText;
    mii.hbmpItem = hBitmap; 

    InsertMenuItemW(hMenu, indexMenu, TRUE, &mii);
    
    // Storing in member to clean up later
    if (hBitmap) m_hMenuBitmap = hBitmap;

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
        RunApp(m_szSelectedFolder);
    }

    return S_OK;
}

IFACEMETHODIMP CContextMenu::GetCommandString(UINT_PTR idCmd, UINT uType, UINT* pReserved, LPSTR pszName, UINT cchMax)
{
    return E_NOTIMPL;
}

void CContextMenu::RunApp(const std::wstring& folderPath)
{
    // 1. Convert path to UTF-8
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, folderPath.c_str(), -1, NULL, 0, NULL, NULL);
    std::string strPath(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, folderPath.c_str(), -1, &strPath[0], size_needed, NULL, NULL);
    // Remove null terminator if present at end for cleaner raw byte send
    if (strPath.length() > 0 && strPath.back() == '\0') {
        strPath.pop_back();
    }

    // 2. Try to connect to existing server via Named Pipe
    // QLocalServer uses "\\.\pipe\" + serverName
    const wchar_t* pipeName = L"\\\\.\\pipe\\ExplorerSelectorServer";
    
    HANDLE hPipe = CreateFileW(pipeName, GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (hPipe != INVALID_HANDLE_VALUE) {
        DWORD written;
        // Write payload
        WriteFile(hPipe, strPath.c_str(), (DWORD)strPath.length(), &written, NULL);
        CloseHandle(hPipe);
        return;
    }

    // 3. Fallback: Start Process
    // Get the DLL path to find the EXE
    wchar_t szDllPath[MAX_PATH];
    GetModuleFileNameW(g_hInst, szDllPath, MAX_PATH);
    PathRemoveFileSpecW(szDllPath);
    
    std::wstring exePath = std::wstring(szDllPath) + L"\\ExplorerSelector.exe";

    // Construct command line: "PathToExe" "FolderPath"
    std::wstring commandLine = L"\"" + exePath + L"\" \"" + folderPath + L"\"";

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
}

// IExplorerCommand Implementation

IFACEMETHODIMP CContextMenu::GetTitle(IShellItemArray* psiItemArray, LPWSTR* ppszName)
{
    wchar_t szMenuText[128];
    if (0 == LoadStringW(g_hInst, IDS_MENU_TEXT, szMenuText, ARRAYSIZE(szMenuText))) {
        wcscpy_s(szMenuText, L"Find Files");
    }
    return SHStrDupW(szMenuText, ppszName);
}

IFACEMETHODIMP CContextMenu::GetIcon(IShellItemArray* psiItemArray, LPWSTR* ppszIcon)
{
    // Returns "PathToDll,-IconID"
    wchar_t szDllPath[MAX_PATH];
    GetModuleFileNameW(g_hInst, szDllPath, MAX_PATH);
    std::wstring iconPath = std::wstring(szDllPath) + L",-" + std::to_wstring(IDI_ICON1);
    return SHStrDupW(iconPath.c_str(), ppszIcon);
}

IFACEMETHODIMP CContextMenu::GetToolTip(IShellItemArray* psiItemArray, LPWSTR* ppszInfotip)
{
    return E_NOTIMPL;
}

IFACEMETHODIMP CContextMenu::GetCanonicalName(GUID* pguidCommandName)
{
    *pguidCommandName = CLSID_ExplorerSelector;
    return S_OK;
}

IFACEMETHODIMP CContextMenu::GetState(IShellItemArray* psiItemArray, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState)
{
    *pCmdState = ECS_ENABLED;
    return S_OK;
}

IFACEMETHODIMP CContextMenu::Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc)
{
    if (!psiItemArray) return E_INVALIDARG;

    DWORD count = 0;
    psiItemArray->GetCount(&count);
    
    if (count > 0)
    {
        IShellItem* psi = NULL;
        if (SUCCEEDED(psiItemArray->GetItemAt(0, &psi)))
        {
            LPWSTR pszPath = NULL;
            if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &pszPath)))
            {
                RunApp(pszPath);
                CoTaskMemFree(pszPath);
            }
            psi->Release();
        }
    }
    return S_OK;
}

IFACEMETHODIMP CContextMenu::GetFlags(EXPCMDFLAGS* pFlags)
{
    *pFlags = ECF_DEFAULT;
    return S_OK;
}

IFACEMETHODIMP CContextMenu::EnumSubCommands(IEnumExplorerCommand** ppEnum)
{
    *ppEnum = NULL;
    return E_NOTIMPL;
}
