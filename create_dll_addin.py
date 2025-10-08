import sys
import os
import subprocess
import tempfile

def create_dll_addin():
    """יוצר תוסף COM עם קובץ DLL אמיתי"""
    
    # קוד C++ שיוצר DLL אמיתי
    cpp_code = '''
#include <windows.h>
#include <comdef.h>
#include <msoutl.h>

// GUID ייחודי
const GUID CLSID_OutlookAddin = {0xEEEEEEEE, 0xEEEE, 0xEEEE, {0xEE, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE, 0xEE}};

// ממשק התוסף
class COutlookAddin : public IDispatch
{
private:
    LONG m_cRef;
    
public:
    COutlookAddin() : m_cRef(1) 
    {
        // יצירת קובץ טסט
        HANDLE hFile = CreateFileA("C:\\\\Users\\\\ronni\\\\AppData\\\\Local\\\\Temp\\\\dll_addin_loaded.txt",
                                   GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE) {
            const char* message = "DLL Addin Loaded Successfully!";
            DWORD bytesWritten;
            WriteFile(hFile, message, strlen(message), &bytesWritten, NULL);
            CloseHandle(hFile);
        }
    }
    
    ~COutlookAddin() 
    {
        // יצירת קובץ טסט נוסף
        HANDLE hFile = CreateFileA("C:\\\\Users\\\\ronni\\\\AppData\\\\Local\\\\Temp\\\\dll_addin_destroyed.txt",
                                   GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE) {
            const char* message = "DLL Addin Destroyed Successfully!";
            DWORD bytesWritten;
            WriteFile(hFile, message, strlen(message), &bytesWritten, NULL);
            CloseHandle(hFile);
        }
    }
    
    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, void** ppv) override
    {
        if (riid == IID_IUnknown || riid == IID_IDispatch) {
            *ppv = static_cast<IDispatch*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    
    STDMETHOD_(ULONG, AddRef)() override
    {
        return InterlockedIncrement(&m_cRef);
    }
    
    STDMETHOD_(ULONG, Release)() override
    {
        LONG cRef = InterlockedDecrement(&m_cRef);
        if (cRef == 0) {
            delete this;
        }
        return cRef;
    }
    
    // IDispatch methods
    STDMETHOD(GetTypeInfoCount)(UINT* pctinfo) override
    {
        *pctinfo = 0;
        return S_OK;
    }
    
    STDMETHOD(GetTypeInfo)(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) override
    {
        return E_NOTIMPL;
    }
    
    STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) override
    {
        return E_NOTIMPL;
    }
    
    STDMETHOD(Invoke)(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr) override
    {
        return S_OK;
    }
};

// Factory class
class CAddinFactory : public IClassFactory
{
private:
    LONG m_cRef;
    
public:
    CAddinFactory() : m_cRef(1) {}
    
    // IUnknown methods
    STDMETHOD(QueryInterface)(REFIID riid, void** ppv) override
    {
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *ppv = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    
    STDMETHOD_(ULONG, AddRef)() override
    {
        return InterlockedIncrement(&m_cRef);
    }
    
    STDMETHOD_(ULONG, Release)() override
    {
        LONG cRef = InterlockedDecrement(&m_cRef);
        if (cRef == 0) {
            delete this;
        }
        return cRef;
    }
    
    // IClassFactory methods
    STDMETHOD(CreateInstance)(IUnknown* pUnkOuter, REFIID riid, void** ppv) override
    {
        if (pUnkOuter != NULL) {
            return CLASS_E_NOAGGREGATION;
        }
        
        COutlookAddin* pAddin = new COutlookAddin();
        if (pAddin == NULL) {
            return E_OUTOFMEMORY;
        }
        
        HRESULT hr = pAddin->QueryInterface(riid, ppv);
        pAddin->Release();
        return hr;
    }
    
    STDMETHOD(LockServer)(BOOL fLock) override
    {
        return S_OK;
    }
};

// DLL exports
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    if (rclsid == CLSID_OutlookAddin) {
        CAddinFactory* pFactory = new CAddinFactory();
        if (pFactory == NULL) {
            return E_OUTOFMEMORY;
        }
        
        HRESULT hr = pFactory->QueryInterface(riid, ppv);
        pFactory->Release();
        return hr;
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow()
{
    return S_FALSE;
}

STDAPI DllRegisterServer()
{
    return S_OK;
}

STDAPI DllUnregisterServer()
{
    return S_OK;
}

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
    switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
        // יצירת קובץ טסט
        HANDLE hFile = CreateFileA("C:\\\\Users\\\\ronni\\\\AppData\\\\Local\\\\Temp\\\\dll_main_loaded.txt",
                                   GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE) {
            const char* message = "DLL Main Loaded Successfully!";
            DWORD bytesWritten;
            WriteFile(hFile, message, strlen(message), &bytesWritten, NULL);
            CloseHandle(hFile);
        }
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
'''
    
    # כתיבת הקוד לקובץ
    with open('outlook_addin.cpp', 'w', encoding='utf-8') as f:
        f.write(cpp_code)
    
    print("Created outlook_addin.cpp")
    return True

if __name__ == '__main__':
    print("Creating DLL Addin...")
    create_dll_addin()
    print("DLL Addin source created!")


