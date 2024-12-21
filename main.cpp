#include <windows.h>
#include <string>
#include <sstream>
#include <wbemidl.h>
#include <comdef.h>
#include <commctrl.h>
#include <vector>
#include <d3d9.h>


#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "d3d9.lib")

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

int GetLogicalCoreCount() {
    SYSTEM_INFO sysInfo;
    GetSystemInfo(&sysInfo);
    return sysInfo.dwNumberOfProcessors;
}
BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData) {
    std::wstringstream* ss = reinterpret_cast<std::wstringstream*>(dwData);

    MONITORINFOEX monitorInfo;
    monitorInfo.cbSize = sizeof(MONITORINFOEX);

    if (GetMonitorInfo(hMonitor, &monitorInfo)) {
        *ss << L"Monitor:";
        *ss << L"  Device Name: " << monitorInfo.szDevice << L"\r\n";
        *ss << L"  Resolution: " << (monitorInfo.rcMonitor.right - monitorInfo.rcMonitor.left)
            << L"x" << (monitorInfo.rcMonitor.bottom - monitorInfo.rcMonitor.top) << L" pixels\r\n";

        if (monitorInfo.dwFlags & MONITORINFOF_PRIMARY) {
            *ss << L"  Primary Monitor: Yes\r\n";
        }
        else {
            *ss << L"  Primary Monitor: No\r\n";
        }
        *ss << L"\r\n";
    }

    return TRUE; 
}
//вивід інформації про ос
std::wstring GetOSInfo() {
    std::wstringstream ss;
    ss << L"=== Operating System Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    
    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"),
        bstr_t("SELECT Caption, Version, OSArchitecture, BuildNumber, InstallDate, CSName, RegisteredUser FROM Win32_OperatingSystem"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL, &pEnumerator);
    if (FAILED(hres)) {
        ss << L"WMI query for OS failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"Caption", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"OS Name: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Version: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"OSArchitecture", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Architecture: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"BuildNumber", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Build Number: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"InstallDate", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            std::wstring dateStr = vtProp.bstrVal;
            if (dateStr.length() >= 8) {
                ss << L"Install Date: " << dateStr.substr(6, 2) << L"." << dateStr.substr(4, 2) << L"." << dateStr.substr(0, 4) << L"\r\n";
            }
            else {
                ss << L"Install Date: " << dateStr << L"\r\n";
            }
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"CSName", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Computer Name: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"RegisteredUser", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Register user: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про процесор
std::wstring GetCPUInfo() {
    std::wstringstream ss;
    ss << L"=== Processor Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    
    hres = CoInitializeSecurity(
        NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL
    );
    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(
        CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc
    );
    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc
    );
    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT Name, NumberOfCores, MaxClockSpeed FROM Win32_Processor"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator
    );

    if (FAILED(hres)) {
        ss << L"WMI query failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) {
            break;
        }

        VARIANT vtProp;

        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) { 
            ss << L"Processor Name: " << vtProp.bstrVal << L"\r\n";
        }
        else {
            ss << L"Processor Name: N/A\r\n";
        }
        VariantClear(&vtProp);

        int logicalCores = GetLogicalCoreCount();
        ss << L"Logical Processors: " << logicalCores << L"\r\n";

        hr = pclsObj->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_I4) {  
            ss << L"Number of Cores: " << vtProp.intVal << L"\r\n";
        }
        else {
            ss << L"Number of Cores: N/A\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"MaxClockSpeed", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_UI4) { 
            ss << L"Original Clock Speed (MHz): " << vtProp.uintVal << L"\r\n";
        }
        else {
            ss << L"Original Clock Speed (MHz): N/A\r\n";
        }
        VariantClear(&vtProp);

    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про озу
std::wstring GetMemoryInfo() {
    std::wstringstream ss;
    ss << L"=== Memory Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    
    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);

    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);

    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    // Выполнение запроса WMI
    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"), bstr_t("SELECT TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);

    if (FAILED(hres)) {
        ss << L"WMI query failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        // Общий объём оперативной памяти
        hr = pclsObj->Get(L"TotalVisibleMemorySize", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ULONGLONG totalMemoryKB = _wtoi64(vtProp.bstrVal);
            ss << L"Total Memory: " << (totalMemoryKB / 1024) << L" MB\r\n";
        }
        VariantClear(&vtProp);

        // Свободная память
        hr = pclsObj->Get(L"FreePhysicalMemory", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ULONGLONG freeMemoryKB = _wtoi64(vtProp.bstrVal);
            ss << L"Free Memory: " << (freeMemoryKB / 1024) << L" MB\r\n";
        }
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    // Очистка
    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про BIOS
std::wstring GetBIOSInfo() {
    std::wstringstream ss;
    ss << L"=== BIOS Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"), bstr_t("SELECT * FROM Win32_BIOS"), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
    if (FAILED(hres)) {
        ss << L"WMI query for BIOS failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Type BIOS: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"SMBIOSBIOSVersion", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"BIOS Version: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"ReleaseDate", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            std::wstring dateStr = vtProp.bstrVal;
            if (dateStr.length() >= 8) {
                ss << L"System BIOS Date: " << dateStr.substr(6, 2) << L"." << dateStr.substr(4, 2) << L"." << dateStr.substr(0, 4) << L"\r\n";
            } else {
                ss << L"System BIOS Date: " << dateStr << L"\r\n";
            }
        }
        VariantClear(&vtProp);
    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про системну плату
std::wstring GetMotherboardInfo() {
    std::wstringstream ss;
    ss << L"=== Motherboard Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    
    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"), bstr_t("SELECT Product, Manufacturer, SerialNumber, Version FROM Win32_BaseBoard"), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
    if (FAILED(hres)) {
        ss << L"WMI query for motherboard failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"Product", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Motherboard Name: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Manufacturer: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    hres = pSvc->ExecQuery(bstr_t("WQL"), bstr_t("SELECT MaxClockSpeed, ExtClock FROM Win32_Processor"), WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);
    if (FAILED(hres)) {
        ss << L"WMI query for processor failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        // Частота шины
        hr = pclsObj->Get(L"ExtClock", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_I4) {
            ss << L"Bus Speed: " << vtProp.intVal << L" MHz\r\n";
        }
        VariantClear(&vtProp);


        pclsObj->Release();
    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про дисках
std::wstring GetDiskInfo() {
    std::wstringstream ss;
    ss << L"=== Disk Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);

    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);

    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"), bstr_t("SELECT DeviceID, FreeSpace, Size, FileSystem FROM Win32_LogicalDisk WHERE DriveType = 3"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pEnumerator);

    if (FAILED(hres)) {
        ss << L"WMI query for disks failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"DeviceID", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) { 
            ss << L"Drive: " << vtProp.bstrVal << L"\r\n";
        }
        else {
            ss << L"Drive: N/A\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"FreeSpace", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && (vtProp.vt == VT_UI8 || SUCCEEDED(VariantChangeType(&vtProp, &vtProp, 0, VT_UI8)))) {
            ss << L"Free Space: " << (vtProp.ullVal / (1024 * 1024 * 1024)) << L" GB\r\n";
        }
        else {
            ss << L"Free Space: N/A\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"Size", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && (vtProp.vt == VT_UI8 || SUCCEEDED(VariantChangeType(&vtProp, &vtProp, 0, VT_UI8)))) {
            ss << L"Total Size: " << (vtProp.ullVal / (1024 * 1024 * 1024)) << L" GB\r\n";
        }
        else {
            ss << L"Total Size: N/A\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"FileSystem", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) { 
            ss << L"File System: " << vtProp.bstrVal << L"\r\n";
        }
        else {
            ss << L"File System: N/A\r\n";
        }
        VariantClear(&vtProp);
        ss << L"\r\n";

        pclsObj->Release();
    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про монитори
std::wstring GetMonitorInfo() {
    std::wstringstream ss;
    ss << L"=== Monitor Information ===\r\n";

    EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, reinterpret_cast<LPARAM>(&ss));

    return ss.str();
}
//вивід інформації про видеоадаптер
std::wstring GetVideoAdapterInfo() {
    std::wstringstream ss;
    ss << L"=== Video Adapter Information ===\r\n";

    IDirect3D9* pD3D = Direct3DCreate9(D3D_SDK_VERSION);
    if (pD3D) {
        D3DADAPTER_IDENTIFIER9 adapterIdentifier;
        D3DCAPS9 caps;

        for (UINT i = 0; i < pD3D->GetAdapterCount(); ++i) {
            if (SUCCEEDED(pD3D->GetAdapterIdentifier(i, 0, &adapterIdentifier)) &&
                SUCCEEDED(pD3D->GetDeviceCaps(i, D3DDEVTYPE_HAL, &caps))) {
                ss << L"Adapter " << i << L":\r\n";
                ss << L"  Name: " << adapterIdentifier.Description << L"\r\n";
                ss << L"  Driver Version: " << adapterIdentifier.DriverVersion.HighPart << L"." << adapterIdentifier.DriverVersion.LowPart << L"\r\n";
                ss << L"  Video Memory: " << (caps.MaxTextureWidth * caps.MaxTextureHeight * 4 / (1024 * 1024)) << L" MB\r\n";
            }
            ss << L"\r\n";
        }
        
        pD3D->Release();
    }
    else {
        ss << L"Failed to create Direct3D object.\r\n";
    }

    return ss.str();
}
//вивід інформації про мережу
std::wstring GetNetworkInfo() {
    std::wstringstream ss;
    ss << L"=== Network Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"),
        bstr_t("SELECT Description, IPAddress, MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL, &pEnumerator);
    if (FAILED(hres)) {
        ss << L"WMI query for network failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"Description", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Adapter: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"MACAddress", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"MAC Address: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        hr = pclsObj->Get(L"IPAddress", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == (VT_ARRAY | VT_BSTR)) {
            SAFEARRAY* psa = vtProp.parray;
            BSTR* pbstr;
            SafeArrayAccessData(psa, (void**)&pbstr);

            ss << L"IP Address: " << pbstr[0] << L"\r\n";
            SafeArrayUnaccessData(psa);
            
        }
        ss << L"\r\n";
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про переферію
std::wstring GetPeripheralInfo() {
    std::wstringstream ss;
    ss << L"=== Peripheral Devices Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"),
        bstr_t("SELECT Name, Description FROM Win32_PnPEntity"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL, &pEnumerator);
    if (FAILED(hres)) {
        ss << L"WMI query for peripherals failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            ss << L"Device: " << vtProp.bstrVal << L"\r\n";
        }
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}
//вивід інформації про usb пристрої
std::wstring GetUSBDevicesInfo() {
    std::wstringstream ss;
    ss << L"=== USB Devices Information ===\r\n";

    HRESULT hres;
    IWbemLocator* pLoc = NULL;
    IWbemServices* pSvc = NULL;

    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        ss << L"COM initialization failed.\r\n";
        return ss.str();
    }

    hres = CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
    if (FAILED(hres)) {
        ss << L"COM security initialization failed.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);
    if (FAILED(hres)) {
        ss << L"Failed to create IWbemLocator.\r\n";
        CoUninitialize();
        return ss.str();
    }

    hres = pLoc->ConnectServer(_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0, NULL, 0, 0, &pSvc);
    if (FAILED(hres)) {
        ss << L"Could not connect to WMI namespace.\r\n";
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_USBControllerDevice"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL, &pEnumerator);
    if (FAILED(hres)) {
        ss << L"WMI query for USB devices failed.\r\n";
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        return ss.str();
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;

    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;

        hr = pclsObj->Get(L"Dependent", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr) && vtProp.vt == VT_BSTR) {
            std::wstring devicePath = vtProp.bstrVal;
            size_t pos = devicePath.find(L"DeviceID=\"");
            if (pos != std::wstring::npos) {
                std::wstring deviceId = devicePath.substr(pos + 10);
                deviceId = deviceId.substr(0, deviceId.length() - 1);

                IEnumWbemClassObject* pDeviceEnumerator = NULL;
                std::wstring query = L"SELECT Name, Description FROM Win32_PnPEntity WHERE DeviceID=\"" + deviceId + L"\"";
                hres = pSvc->ExecQuery(bstr_t("WQL"), bstr_t(query.c_str()),
                    WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY, NULL, &pDeviceEnumerator);
                if (SUCCEEDED(hres)) {
                    IWbemClassObject* pDeviceObj = NULL;
                    ULONG uDeviceReturn = 0;

                    while (pDeviceEnumerator) {
                        HRESULT hrDevice = pDeviceEnumerator->Next(WBEM_INFINITE, 1, &pDeviceObj, &uDeviceReturn);
                        if (0 == uDeviceReturn) break;

                        VARIANT vtName, vtDescription;

                        hrDevice = pDeviceObj->Get(L"Name", 0, &vtName, 0, 0);
                        if (SUCCEEDED(hrDevice) && vtName.vt == VT_BSTR) {
                            ss << L"Device Name: " << vtName.bstrVal << L"\r\n";
                        }
                        VariantClear(&vtName);

                        hrDevice = pDeviceObj->Get(L"Description", 0, &vtDescription, 0, 0);
                        if (SUCCEEDED(hrDevice) && vtDescription.vt == VT_BSTR) {
                            ss << L"Description: " << vtDescription.bstrVal << L"\r\n";
                        }
                        VariantClear(&vtDescription);

                        pDeviceObj->Release();
                    }
                    pDeviceEnumerator->Release();
                }
            }
        }
        VariantClear(&vtProp);

        pclsObj->Release();
    }

    pEnumerator->Release();
    pSvc->Release();
    pLoc->Release();
    CoUninitialize();

    return ss.str();
}

HWND hOutput;
HWND hScrollBar;

void UpdateOutput(const std::wstring& text) {
    SetWindowTextW(hOutput, text.c_str());
}


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"MainWindowClass";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    RegisterClass(&wc);

    HWND hwnd = CreateWindowEx(
        0,
        wc.lpszClassName,
        L"System Information Viewer",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
        NULL, NULL, hInstance, NULL
    );

    if (!hwnd) return 0;

    ShowWindow(hwnd, nCmdShow);

    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    static HWND hMenu;
    static HWND hDivider;

    switch (uMsg) {
    case WM_CREATE: {
        hMenu = CreateWindowEx(
            0, L"LISTBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY,
            0, 0, 200, 600,
            hwnd, (HMENU)1, NULL, NULL
        );


        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"General Info: OS");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"CPU");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"Memory");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"BIOS");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"Motherboard");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"Disks");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"Devices: Monitor");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"VideoAdapter");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"Network");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"Devices: Peripherals");
        SendMessage(hMenu, LB_ADDSTRING, 0, (LPARAM)L"USB Devices");

        hDivider = CreateWindowEx(
    0, L"STATIC", NULL,
    WS_CHILD | WS_VISIBLE | SS_ETCHEDVERT,
    200, 0, 2, 600,
    hwnd, NULL, NULL, NULL
);

        hOutput = CreateWindowEx(
            WS_EX_CLIENTEDGE, L"EDIT", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL,
            202, 0, 598, 600,
            hwnd, NULL, NULL, NULL
        );

        return 0;
    }
    case WM_COMMAND: {
        if (HIWORD(wParam) == LBN_SELCHANGE) {
            int sel = SendMessage(hMenu, LB_GETCURSEL, 0, 0);
            std::wstring result;

            wchar_t debugMessage[100];
            swprintf_s(debugMessage, L"Selected menu item: %d\n", sel);
            OutputDebugStringW(debugMessage);

            switch (sel) {
            case 0: result = GetOSInfo(); break;
            case 1: result = GetCPUInfo(); break;
            case 2: result = GetMemoryInfo(); break;
            case 3: result = GetBIOSInfo(); break;
            case 4: result = GetMotherboardInfo(); break;
            case 5: result = GetDiskInfo(); break;
            case 6: result = GetMonitorInfo(); break;
            case 7: result = GetVideoAdapterInfo(); break;
            case 8: result = GetNetworkInfo(); break;
            case 9: result = GetPeripheralInfo(); break;
            case 10: result = GetUSBDevicesInfo(); break;
            default: result = L"Unknown selection."; break;
            }

            OutputDebugStringW(result.c_str());
            UpdateOutput(result);
        }
        return 0;
    }


    case WM_SIZE: {
        RECT rc;
        GetClientRect(hwnd, &rc);
        SetWindowPos(hMenu, NULL, 0, 0, 200, rc.bottom, SWP_NOZORDER);
        SetWindowPos(hDivider, NULL, 200, 0, 2, rc.bottom, SWP_NOZORDER);
        SetWindowPos(hOutput, NULL, 202, 0, rc.right - 202, rc.bottom, SWP_NOZORDER);
        return 0;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

