#include "pch.h"
#include "ClassFactory.h"
#include <strsafe.h>
#include <new>

HINSTANCE g_hInst = NULL;
long g_cDllRef = 0;

BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        DisableThreadLibraryCalls(hModule);
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (IsEqualCLSID(CLSID_ExplorerSelector, rclsid))
    {
        CClassFactory* pClassFactory = new (std::nothrow) CClassFactory();
        if (NULL == pClassFactory)
        {
            return E_OUTOFMEMORY;
        }
        HRESULT hr = pClassFactory->QueryInterface(riid, ppv);
        pClassFactory->Release();
        return hr;
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow()
{
    return g_cDllRef > 0 ? S_FALSE : S_OK;
}

HRESULT RegisterInHKCR(PCWSTR pszKey, PCWSTR pszSubkey, PCWSTR pszValue)
{
    HKEY hKey;
    HRESULT hr = E_FAIL;
    std::wstring strKey = pszKey;
    if (pszSubkey)
    {
        strKey += L"\\";
        strKey += pszSubkey;
    }

    if (ERROR_SUCCESS == RegCreateKeyExW(HKEY_CLASSES_ROOT, strKey.c_str(), 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL))
    {
        if (pszValue)
        {
            if (ERROR_SUCCESS == RegSetValueExW(hKey, NULL, 0, REG_SZ, (const BYTE*)pszValue, (lstrlenW(pszValue) + 1) * sizeof(wchar_t)))
            {
                hr = S_OK;
            }
        }
        else
        {
            hr = S_OK;
        }
        RegCloseKey(hKey);
    }
    return hr;
}

STDAPI DllRegisterServer()
{
    // Register CLSID
    std::wstring strCLSID = L"CLSID\\" + std::wstring(CLSID_ExplorerSelector_String);
    HRESULT hr = RegisterInHKCR(strCLSID.c_str(), NULL, L"ExplorerSelector Context Menu");
    if (SUCCEEDED(hr))
    {
        wchar_t szModule[MAX_PATH];
        if (GetModuleFileNameW(g_hInst, szModule, ARRAYSIZE(szModule)))
        {
            hr = RegisterInHKCR(strCLSID.c_str(), L"InprocServer32", szModule);
            if (SUCCEEDED(hr))
            {
                hr = RegisterInHKCR(strCLSID.c_str(), L"InprocServer32\\ThreadingModel", L"Apartment");
            }
        }
    }

    // Register Shell Extension
    if (SUCCEEDED(hr))
    {
        // Directory\Background (Right click on empty space in folder)
        // Legacy (Windows 10 "Show more options" / Old Windows)
        hr = RegisterInHKCR(L"Directory\\Background\\shellex\\ContextMenuHandlers", L"ExplorerSelector", CLSID_ExplorerSelector_String);
        
        if (SUCCEEDED(hr))
        {
            // Modern Windows 11 Registration (ExplorerCommandHandler)
            // Note: For this to appear in the top-level menu on Windows 11, the app usually needs to be packaged (Sparse Package)
            // or registered via identity. However, registering the handler is the first step.
            hr = RegisterInHKCR(L"Directory\\Background\\shell\\ExplorerSelector\\ExplorerCommandHandler", NULL, CLSID_ExplorerSelector_String);
        }
    }

    if (SUCCEEDED(hr))
    {
        SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
    }

    return hr;
}

STDAPI DllUnregisterServer()
{
    // Unregister CLSID
    std::wstring strCLSID = L"CLSID\\" + std::wstring(CLSID_ExplorerSelector_String);
    RegDeleteTreeW(HKEY_CLASSES_ROOT, strCLSID.c_str());

    // Unregister Shell Extension
    RegDeleteTreeW(HKEY_CLASSES_ROOT, L"Directory\\Background\\shellex\\ContextMenuHandlers\\ExplorerSelector");
    RegDeleteTreeW(HKEY_CLASSES_ROOT, L"Directory\\Background\\shell\\ExplorerSelector");

    return S_OK;
}