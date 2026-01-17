#include "pch.h"
#include "ClassFactory.h"
#include "ContextMenu.h"
#include <new>

extern long g_cDllRef;

CClassFactory::CClassFactory() : m_cRef(1)
{
    InterlockedIncrement(&g_cDllRef);
}

CClassFactory::~CClassFactory()
{
    InterlockedDecrement(&g_cDllRef);
}

// IUnknown
IFACEMETHODIMP CClassFactory::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(CClassFactory, IClassFactory),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) CClassFactory::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) CClassFactory::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }
    return cRef;
}

// IClassFactory
IFACEMETHODIMP CClassFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv)
{
    if (pUnkOuter != NULL)
    {
        return CLASS_E_NOAGGREGATION;
    }

    CContextMenu* pContextMenu = new (std::nothrow) CContextMenu();
    if (NULL == pContextMenu)
    {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = pContextMenu->QueryInterface(riid, ppv);
    pContextMenu->Release();
    return hr;
}

IFACEMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
    if (fLock)
    {
        InterlockedIncrement(&g_cDllRef);
    }
    else
    {
        InterlockedDecrement(&g_cDllRef);
    }
    return S_OK;
}
