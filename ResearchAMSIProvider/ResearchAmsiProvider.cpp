#include "ResearchAmsiProvider.h"

using namespace Microsoft::WRL;

// CLSID {4CEFB33D-4F98-4EA9-9511-CBCC475683F4}

// Define a trace logging provider: {281F855D-5FEA-4E08-8301-FDCA71E23ED3}
TRACELOGGING_DEFINE_PROVIDER(g_traceLoggingProvider, "ResearchAMSIProvider", 
    (0x281f855d, 0x5fea, 0x4e08, 0x83, 0x1, 0xfd, 0xca, 0x71, 0xe2, 0x3e, 0xd3));

HMODULE g_hModule = NULL;

BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved) {
	switch (dwReason) {
		case DLL_PROCESS_ATTACH:
			g_hModule = hModule;
            DisableThreadLibraryCalls(hModule);
            TraceLoggingRegister(g_traceLoggingProvider);
            TraceLoggingWrite(g_traceLoggingProvider, "Loaded");
            Module<InProc>::GetModule().Create();
			break;
		case DLL_PROCESS_DETACH:
            DisableThreadLibraryCalls(hModule);
            TraceLoggingRegister(g_traceLoggingProvider);
            TraceLoggingWrite(g_traceLoggingProvider, "Loaded");
            Module<InProc>::GetModule().Create();
			break;
	}
	return TRUE;
}

#pragma region COM server boilerplate
HRESULT WINAPI DllCanUnloadNow()
{
    return Module<InProc>::GetModule().Terminate() ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
    return Module<InProc>::GetModule().GetClassObject(rclsid, riid, ppv);
}
#pragma endregion

class
    DECLSPEC_UUID("4CEFB33D-4F98-4EA9-9511-CBCC475683F4")
    ResearchAMSIProvider : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IAntimalwareProvider, FtmBase>
{
public:
    IFACEMETHOD(Scan)(_In_ IAmsiStream * stream, _Out_ AMSI_RESULT * result) override;
    IFACEMETHOD_(void, CloseSession)(_In_ ULONGLONG session) override;
    IFACEMETHOD(DisplayName)(_Outptr_ LPWSTR * displayName) override;

private:
    // We assign each Scan request a unique number for logging purposes.
    LONG m_requestNumber = 0;
};

template<typename T>
class HeapMemPtr
{
public:
    HeapMemPtr() { }
    HeapMemPtr(const HeapMemPtr& other) = delete;
    HeapMemPtr(HeapMemPtr&& other) : p(other.p) { other.p = nullptr; }
    HeapMemPtr& operator=(const HeapMemPtr& other) = delete;
    HeapMemPtr& operator=(HeapMemPtr&& other) {
        auto t = p; p = other.p; other.p = t;
    }

    ~HeapMemPtr()
    {
        if (p) HeapFree(GetProcessHeap(), 0, p);
    }

    HRESULT Alloc(size_t size)
    {
        p = reinterpret_cast<T*>(HeapAlloc(GetProcessHeap(), 0, size));
        return p ? S_OK : E_OUTOFMEMORY;
    }

    T* Get() { return p; }
    operator bool() { return p != nullptr; }

private:
    T* p = nullptr;
};

template<typename T>
T GetFixedSizeAttribute(_In_ IAmsiStream* stream, _In_ AMSI_ATTRIBUTE attribute)
{
    T result;

    ULONG actualSize;
    if (SUCCEEDED(stream->GetAttribute(attribute, sizeof(T), reinterpret_cast<PBYTE>(&result), &actualSize)) &&
        actualSize == sizeof(T))
    {
        return result;
    }
    return T();
}

HeapMemPtr<wchar_t> GetStringAttribute(_In_ IAmsiStream* stream, _In_ AMSI_ATTRIBUTE attribute)
{
    HeapMemPtr<wchar_t> result;

    ULONG allocSize;
    ULONG actualSize;
    if (stream->GetAttribute(attribute, 0, nullptr, &allocSize) == E_NOT_SUFFICIENT_BUFFER &&
        SUCCEEDED(result.Alloc(allocSize)) &&
        SUCCEEDED(stream->GetAttribute(attribute, allocSize, reinterpret_cast<PBYTE>(result.Get()), &actualSize)) &&
        actualSize <= allocSize)
    {
        return result;
    }
    return HeapMemPtr<wchar_t>();
}

HRESULT ResearchAMSIProvider::Scan(_In_ IAmsiStream* stream, _Out_ AMSI_RESULT* result) {

    LONG requestNumber = InterlockedIncrement(&m_requestNumber);
    TraceLoggingWrite(g_traceLoggingProvider, "Scan Start", TraceLoggingValue(requestNumber));

    HANDLE hFile = NULL;
    BOOL bSuccess = TRUE;

    auto appName = GetStringAttribute(stream, AMSI_ATTRIBUTE_APP_NAME);
    auto contentName = GetStringAttribute(stream, AMSI_ATTRIBUTE_CONTENT_NAME);
    auto contentSize = GetFixedSizeAttribute<ULONGLONG>(stream, AMSI_ATTRIBUTE_CONTENT_SIZE);
    auto session = GetFixedSizeAttribute<ULONGLONG>(stream, AMSI_ATTRIBUTE_SESSION);
    auto contentAddress = GetFixedSizeAttribute<PBYTE>(stream, AMSI_ATTRIBUTE_CONTENT_ADDRESS);

    TraceLoggingWrite(g_traceLoggingProvider, "Attributes",
        TraceLoggingValue(requestNumber),
        TraceLoggingWideString(appName.Get(), "App Name"),
        TraceLoggingWideString(contentName.Get(), "Content Name"),
        TraceLoggingUInt64(contentSize, "Content Size"),
        TraceLoggingUInt64(session, "Session"),
        TraceLoggingPointer(contentAddress, "Content Address"));
    
    CHAR lpPathName[MAX_PATH + 1] = { 0 };
    CHAR lpFileName[MAX_PATH] = { 0 };

    DWORD dwRet = GetTempPathA(sizeof(lpPathName), lpPathName);
    if (dwRet == 0) {
        TraceLoggingWrite(g_traceLoggingProvider, "Error Creating Temp Path");
        goto done;
    }
    dwRet = GetTempFileNameA(lpPathName, "SCR", 0, lpFileName);
    if (dwRet == 0) {
        TraceLoggingWrite(g_traceLoggingProvider, "Error Creating Temp File");
        goto done;
    }
    hFile = CreateFileA(lpFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL, NULL);
    bSuccess = WriteFile(hFile, contentAddress, (ULONGLONG)contentSize, NULL, NULL);

done:

    TraceLoggingWrite(g_traceLoggingProvider, "Scan End", TraceLoggingValue(requestNumber));

    *result = AMSI_RESULT_NOT_DETECTED; // Do nothing
    //*result = AMSI_RESULT_DETECTED; // Block all the things

    if (hFile)
        CloseHandle(hFile);
    
    return S_OK;
}

void ResearchAMSIProvider::CloseSession(_In_ ULONGLONG session)
{
    TraceLoggingWrite(g_traceLoggingProvider, "Close session",
        TraceLoggingValue(session));
}

HRESULT ResearchAMSIProvider::DisplayName(_Outptr_ LPWSTR* displayName)
{
    *displayName = const_cast<LPWSTR>(L"Sample AMSI Provider");
    return S_OK;
}

CoCreatableClass(ResearchAMSIProvider);

#pragma region Install / uninstall
HRESULT SetKeyStringValue(_In_ HKEY key, _In_opt_ PCWSTR subkey, _In_opt_ PCWSTR valueName, _In_ PCWSTR stringValue)
{
    LONG status = RegSetKeyValue(key, subkey, valueName, REG_SZ, stringValue, (wcslen(stringValue) + 1) * sizeof(wchar_t));
    return HRESULT_FROM_WIN32(status);
}

STDAPI DllRegisterServer()
{
    wchar_t modulePath[MAX_PATH];
    if (GetModuleFileName(g_hModule, modulePath, ARRAYSIZE(modulePath)) >= ARRAYSIZE(modulePath))
    {
        return E_UNEXPECTED;
    }

    // Create a standard COM registration for our CLSID.
    // The class must be registered as "Both" threading model
    // and support multithreaded access.
    wchar_t clsidString[40];
    if (StringFromGUID2(__uuidof(ResearchAMSIProvider), clsidString, ARRAYSIZE(clsidString)) == 0)
    {
        return E_UNEXPECTED;
    }

    wchar_t keyPath[200];
    HRESULT hr = StringCchPrintf(keyPath, ARRAYSIZE(keyPath), L"Software\\Classes\\CLSID\\%ls", clsidString);
    if (FAILED(hr)) return hr;

    hr = SetKeyStringValue(HKEY_LOCAL_MACHINE, keyPath, nullptr, L"ResearchAmsiProvider");
    if (FAILED(hr)) return hr;

    hr = StringCchPrintf(keyPath, ARRAYSIZE(keyPath), L"Software\\Classes\\CLSID\\%ls\\InProcServer32", clsidString);
    if (FAILED(hr)) return hr;

    hr = SetKeyStringValue(HKEY_LOCAL_MACHINE, keyPath, nullptr, modulePath);
    if (FAILED(hr)) return hr;

    hr = SetKeyStringValue(HKEY_LOCAL_MACHINE, keyPath, L"ThreadingModel", L"Both");
    if (FAILED(hr)) return hr;

    // Register this CLSID as an anti-malware provider.
    hr = StringCchPrintf(keyPath, ARRAYSIZE(keyPath), L"Software\\Microsoft\\AMSI\\Providers\\%ls", clsidString);
    if (FAILED(hr)) return hr;

    hr = SetKeyStringValue(HKEY_LOCAL_MACHINE, keyPath, nullptr, L"ResearchAmsiProvider");
    if (FAILED(hr)) return hr;

    return S_OK;
}

STDAPI DllUnregisterServer()
{
    wchar_t clsidString[40];
    if (StringFromGUID2(__uuidof(ResearchAMSIProvider), clsidString, ARRAYSIZE(clsidString)) == 0)
    {
        return E_UNEXPECTED;
    }

    // Unregister this CLSID as an anti-malware provider.
    wchar_t keyPath[200];
    HRESULT hr = StringCchPrintf(keyPath, ARRAYSIZE(keyPath), L"Software\\Microsoft\\AMSI\\Providers\\%ls", clsidString);
    if (FAILED(hr)) return hr;
    LONG status = RegDeleteTree(HKEY_LOCAL_MACHINE, keyPath);
    if (status != NO_ERROR && status != ERROR_PATH_NOT_FOUND) return HRESULT_FROM_WIN32(status);

    // Unregister this CLSID as a COM server.
    hr = StringCchPrintf(keyPath, ARRAYSIZE(keyPath), L"Software\\Classes\\CLSID\\%ls", clsidString);
    if (FAILED(hr)) return hr;
    status = RegDeleteTree(HKEY_LOCAL_MACHINE, keyPath);
    if (status != NO_ERROR && status != ERROR_PATH_NOT_FOUND) return HRESULT_FROM_WIN32(status);

    return S_OK;
}
#pragma endregion