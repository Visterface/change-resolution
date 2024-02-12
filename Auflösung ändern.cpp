// Auflösung_ändern.cpp : Diese Datei enthält die Funktion "main". Hier beginnt und endet die Ausführung des Programms.
//
#define _WIN32_DCOM
#include <iostream>
using namespace std;
#include <comdef.h>
#include <Wbemidl.h>
#pragma comment(lib, "wbemuuid.lib")
#include <Windows.h>
#include <map>
#include <vector>

static const UINT32 DpiVals[] = { 100,125,150,175,200,225,250,300,350, 400, 450, 500 };

struct BSTRComparer
{
    bool operator() (const BSTR& lhs, const BSTR& rhs) const
    {
        return wcscmp(lhs, rhs) < 0;
    }
};
bool CompareBSTR(BSTR bstr1, BSTR bstr2) {
    if (SysStringLen(bstr1) != SysStringLen(bstr2)) {
        return false;
    }
    return wcscmp(bstr1, bstr2) == 0;
}

int ui(BSTR mod, int width, int height, int scale) {
    for (;;)
    {
        wcout << "Ihr Modell ist: " << mod << endl;
        cout << "Die Aufloesung wurde auf " << width << " x " << height << " geaendert. Die Skalierung ist nun auf: " << scale << "%" << endl;
        cout << "Druecken Sie eine beliebige Taste um das Programm zu beenden.";
        cin.ignore();
        exit(0);
        break;
    }
}

/*Get default DPI scaling percentage.
The OS recommented value.
*/
int GetRecommendedDPIScaling()
{
    int dpi = 0;
    auto retval = SystemParametersInfo(SPI_GETLOGICALDPIOVERRIDE, 0, (LPVOID)&dpi, 1);

    if (retval != 0)
    {
        int currDPI = DpiVals[dpi * -1];
        return currDPI;
    }

    return -1;
}

void SetDpiScaling(int percentScaleToSet)
{
    int recommendedDpiScale = GetRecommendedDPIScaling();

    if (recommendedDpiScale > 0)
    {
        int index = 0, recIndex = 0, setIndex = 0;
        for (const auto& scale : DpiVals)
        {
            if (recommendedDpiScale == scale)
            {
                recIndex = index;
            }
            if (percentScaleToSet == scale)
            {
                setIndex = index;
            }
            index++;
        }

        int relativeIndex = setIndex - recIndex;
        SystemParametersInfo(SPI_SETLOGICALDPIOVERRIDE, relativeIndex, (LPVOID)0, 1);
    }
}

// Get device model
BSTR getModel() {
    BSTR model{};
    HRESULT hres;

    // Initialize COM.
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres))
    {
        cout << "Failed to initialize COM library. "
            << "Error code = 0x"
            << hex << hres << endl;
        exit(EXIT_FAILURE);              // Program has failed.
    }

    // Initialize 
    hres = CoInitializeSecurity(
        NULL,
        -1,      // COM negotiates service                  
        NULL,    // Authentication services
        NULL,    // Reserved
        RPC_C_AUTHN_LEVEL_DEFAULT,    // authentication
        RPC_C_IMP_LEVEL_IMPERSONATE,  // Impersonation
        NULL,             // Authentication info 
        EOAC_NONE,        // Additional capabilities
        NULL              // Reserved
    );


    if (FAILED(hres))
    {
        cout << "Failed to initialize security. "
            << "Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        exit(EXIT_FAILURE);          // Program has failed.
    }

    // Obtain the initial locator to Windows Management
    // on a particular host computer.
    IWbemLocator* pLoc = 0;

    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator, (LPVOID*)&pLoc);

    if (FAILED(hres))
    {
        cout << "Failed to create IWbemLocator object. "
            << "Error code = 0x"
            << hex << hres << endl;
        CoUninitialize();
        exit(EXIT_FAILURE);       // Program has failed.
    }

    IWbemServices* pSvc = 0;

    // Connect to the root\cimv2 namespace with the
    // current user and obtain pointer pSvc
    // to make IWbemServices calls.

    hres = pLoc->ConnectServer(

        _bstr_t(L"ROOT\\CIMV2"), // WMI namespace
        NULL,                    // User name
        NULL,                    // User password
        0,                       // Locale
        NULL,                    // Security flags                 
        0,                       // Authority       
        0,                       // Context object
        &pSvc                    // IWbemServices proxy
    );

    if (FAILED(hres))
    {
        cout << "Could not connect. Error code = 0x"
            << hex << hres << endl;
        pLoc->Release();
        CoUninitialize();
        exit(EXIT_FAILURE);                // Program has failed.
    }

    // Set the IWbemServices proxy so that impersonation
    // of the user (client) occurs.
    hres = CoSetProxyBlanket(

        pSvc,                         // the proxy to set
        RPC_C_AUTHN_WINNT,            // authentication service
        RPC_C_AUTHZ_NONE,             // authorization service
        NULL,                         // Server principal name
        RPC_C_AUTHN_LEVEL_CALL,       // authentication level
        RPC_C_IMP_LEVEL_IMPERSONATE,  // impersonation level
        NULL,                         // client identity 
        EOAC_NONE                     // proxy capabilities     
    );

    if (FAILED(hres))
    {
        cout << "Could not set proxy blanket. Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        exit(EXIT_FAILURE);               // Program has failed.
    }


    // Use the IWbemServices pointer to make requests of WMI. 
    // Make requests here:

    // For example, query for all the running processes
    IEnumWbemClassObject* pEnumerator = NULL;
    hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t("SELECT * FROM Win32_ComputerSystem"),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        &pEnumerator);

    if (FAILED(hres))
    {
        cout << "Query for processes failed. "
            << "Error code = 0x"
            << hex << hres << endl;
        pSvc->Release();
        pLoc->Release();
        CoUninitialize();
        exit(EXIT_FAILURE);               // Program has failed.
    }
    else
    {
        IWbemClassObject* pclsObj;
        ULONG uReturn = 0;

        while (pEnumerator)
        {
            hres = pEnumerator->Next(WBEM_INFINITE, 1,
                &pclsObj, &uReturn);

            if (0 == uReturn)
            {
                break;
            }

            VARIANT vtProp;

            // Get the value of the Name property
            hres = pclsObj->Get(L"Model", 0, &vtProp, 0, 0);
            model = vtProp.bstrVal;
            VariantClear(&vtProp);

            pclsObj->Release();
            pclsObj = NULL;
        }

    }

    // Cleanup
    // ========

    pSvc->Release();
    pLoc->Release();
    pEnumerator->Release();

    CoUninitialize();

    return model;   // Program successfully completed.
}

//Change resolution from 3:2 to 16:9 and vice versa if a compatible surface device. Else fallback to 1920x1080 or 1920x1280.
void change_res() {
    BSTR model = getModel();
    DEVMODE dm;
    ZeroMemory(&dm, sizeof(dm));
    dm.dmSize = sizeof(dm);
    BSTR go_2 = SysAllocString(L"Surface Go 2");
    BSTR go_3 = SysAllocString(L"Surface Go 3");
    BSTR pro_7 = SysAllocString(L"Surface Pro 7");
    BSTR pro_7plus = SysAllocString(L"Surface Pro 7+");
    BSTR pro_8 = SysAllocString(L"Surface Pro 8");
    BSTR pro_8plus = SysAllocString(L"Surface Pro 8+");
    BSTR pro_9 = SysAllocString(L"Surface Pro 9");
    BSTR test = SysAllocString(L"MS-7A39");


    std::map<BSTR, std::vector<int>, BSTRComparer> m = { {go_2, {1920, 1280, 1920, 1080}}, {go_3, {1920, 1280, 1920, 1080}}
    , { pro_7 , { 2736, 1824, 1920, 1080 }}, { pro_7plus , { 2736, 1824, 1920, 1080 }}, { pro_8 , { 2880, 1920, 1920, 1080 }},
    { pro_8plus , { 2880, 1920, 1920, 1080 }},{ pro_9 , { 2880, 1920, 1920, 1080 }}, {test, {2560,1440,1920,1080 }} };

    auto it = m.find(model);
    if (it != m.end()) {
        SetProcessDPIAware();
        if (GetSystemMetrics(0) == it->second[0] && GetSystemMetrics(1) == it->second[1])
        {

            dm.dmPelsWidth = it->second[2];
            dm.dmPelsHeight = it->second[3];
            dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;
            LONG result = ChangeDisplaySettings(&dm, CDS_UPDATEREGISTRY); // Change the display settings
            if (result == DISP_CHANGE_SUCCESSFUL)
            {
                std::cout << "Display settings changed successfully." << std::endl;
            }
            else
            {
                std::cout << "Failed to change display settings." << std::endl;
            }
            SetDpiScaling(125);
            ui(model, it->second[2], it->second[3], 125);
        }
        else if (GetSystemMetrics(0) == it->second[2] && GetSystemMetrics(1) == it->second[3])
        {
            dm.dmPelsWidth = it->second[0];
            dm.dmPelsHeight = it->second[1];
            dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;
            LONG result = ChangeDisplaySettings(&dm, CDS_UPDATEREGISTRY); // Change the display settings
            if (result == DISP_CHANGE_SUCCESSFUL)
            {
                std::cout << "Display settings changed successfully." << std::endl;
            }
            else
            {
                std::cout << "Failed to change display settings." << std::endl;
            }
            SetDpiScaling(GetRecommendedDPIScaling());
            ui(model, it->second[0], it->second[1], GetRecommendedDPIScaling());
        }
    }
    //Fallback auf die am niedrigsten mögliche Auflösung
    else {
        std::cout << "Model nicht kompatibel, versuche auf niedrigst moegliche Aufloesung zu aendern." << std::endl;
        dm.dmPelsWidth = 1920;
        dm.dmPelsHeight = 1280; // Set the height of the display
        dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;
        LONG result = ChangeDisplaySettings(&dm, CDS_UPDATEREGISTRY); // Change the display settings
        if (result == DISP_CHANGE_SUCCESSFUL)
        {
            std::cout << "Display settings changed successfully." << std::endl;
            SetDpiScaling(GetRecommendedDPIScaling());
            ui(model, 1920, 1280, GetRecommendedDPIScaling());
        }
        else
        {
            dm.dmPelsWidth = 1920;
            dm.dmPelsHeight = 1080; // Set the height of the display
            dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;
            LONG result = ChangeDisplaySettings(&dm, CDS_UPDATEREGISTRY); // Change the display settings
            if (result == DISP_CHANGE_SUCCESSFUL)
            {
                std::cout << "Display settings changed successfully." << std::endl;
                SetDpiScaling(GetRecommendedDPIScaling());
                ui(model, 1920, 1080, 125);
            }
            else
            {
                std::cout << "Failed to change display settings." << std::endl;
            }

        }

    }
}

int main()
{
    change_res();
}

