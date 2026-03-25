#include "windows.h"
#include <string>
#include <vector>
#include <memory>
#include <Mmdeviceapi.h>
#include <Shlobj.h>
#include <atlbase.h>
#include <atlcomcli.h>

#include "PolicyConfig.h"
#include "Propidl.h"
#include "Functiondiscoverykeys_devpkey.h"
#include "aesIMMNotificationClient.h"
#include "Prefs.h"
#include "Settings.h"
#include "resource.h"

using namespace std;

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

inline constexpr int  MAX_LOADSTRING    = 100;
inline constexpr UINT WM_USER_SHELLICON = WM_USER + 1;

// Menu item IDs
inline constexpr UINT ID_MENU_ABOUT    = 100;
inline constexpr UINT ID_MENU_DEVICES  = 120;
inline constexpr UINT ID_MENU_EXIT     = 140;
inline constexpr UINT ID_MENU_SETTINGS = 150;
inline constexpr UINT CYCLE_HOTKEY     = 200;

// Timer ID
inline constexpr UINT TIMER_AUDIO_REFRESH = 1234;

// Hotkey flags applied to every RegisterHotKey call
inline constexpr UINT gHotkeyFlags = MOD_NOREPEAT;

// ---------------------------------------------------------------------------
// Globals
// ---------------------------------------------------------------------------

HINSTANCE      hInst;
NOTIFYICONDATA nidApp;
HICON          hMainIcon;
bool           gbDevicesChanged = false;

WCHAR    gszTitle[MAX_LOADSTRING];
WCHAR    gszWindowClass[MAX_LOADSTRING];
WCHAR    gszApplicationToolTip[MAX_LOADSTRING];
HINSTANCE ghInstance;
HWND     gOpenWindow = nullptr;

CQSESPrefs                        gPrefs;
unique_ptr<CMMNotificationClient> pClient;

// ---------------------------------------------------------------------------
// Forward declarations
// ---------------------------------------------------------------------------

ATOM             MyRegisterClass(HINSTANCE hInstance);
BOOL             InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);
void DoSettingsDialog(HINSTANCE hInst, HWND hWnd);

void    UpdateTrayTooltip();
void    InstallNotificationCallback();
int     EnumerateDevices();
bool    SetDefaultAudioPlaybackDevice(LPCWSTR devID);
bool    GetDeviceIcon(HICON* hIcon);
bool    IsDefaultAudioPlaybackDevice(const wstring& s);
wstring GetDefaultAudioPlaybackDevice();

// ---------------------------------------------------------------------------
// Auto-start helpers
// ---------------------------------------------------------------------------

static void MakeStartupLinkPath(const WCHAR* name, WCHAR* path)
{
    LPWSTR wstrStartupPath = nullptr;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_Startup, 0, nullptr, &wstrStartupPath);
    if (SUCCEEDED(hr))
    {
        wcscpy_s(path, MAX_PATH, wstrStartupPath);
        CoTaskMemFree(wstrStartupPath);
        wcscat_s(path, MAX_PATH, L"\\");
        wcscat_s(path, MAX_PATH, name);
        wcscat_s(path, MAX_PATH, L".lnk");
    }
    else
    {
        CoTaskMemFree(wstrStartupPath);
        *path = L'\0';
    }
}

void EnableAutoStart()
{
    IShellLink*   pShellLink   = nullptr;
    IPersistFile* pPersistFile = nullptr;

    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_ALL,
                                  IID_IShellLink, reinterpret_cast<void**>(&pShellLink));
    if (FAILED(hr)) return;

    hr = pShellLink->QueryInterface(IID_IPersistFile,
                                    reinterpret_cast<void**>(&pPersistFile));
    if (SUCCEEDED(hr))
    {
        WCHAR szExeName[MAX_PATH] = {};
        GetModuleFileNameW(nullptr, szExeName, MAX_PATH);
        pShellLink->SetPath(szExeName);
        pShellLink->SetDescription(gszApplicationToolTip);
        pShellLink->SetIconLocation(szExeName, 0);

        WCHAR szStartupLinkName[MAX_PATH] = {};
        MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
        pPersistFile->Save(szStartupLinkName, TRUE);
        pPersistFile->Release();
    }
    pShellLink->Release();
}

void DisableAutoStart()
{
    WCHAR szStartupLinkName[MAX_PATH] = {};
    MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
    if (*szStartupLinkName != L'\0')
        DeleteFileW(szStartupLinkName);
}

UINT CheckAutoStart()
{
    WCHAR szStartupLinkName[MAX_PATH] = {};
    MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
    HANDLE hFile = CreateFileW(szStartupLinkName, GENERIC_READ, 0, nullptr,
                               OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile != INVALID_HANDLE_VALUE)
    {
        CloseHandle(hFile);
        return 1;
    }
    return 0;
}

// ---------------------------------------------------------------------------
// Hotkey string builder
// ---------------------------------------------------------------------------

static wstring MakeHotkeyString(UINT code, UINT mods)
{
    WCHAR keyname[256] = {};
    int   length       = 0;

    auto appendKey = [&](UINT vk)
    {
        if (length > 0) { keyname[length++] = L'+'; keyname[length] = L'\0'; }
        UINT sc = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
        length += GetKeyNameTextW(static_cast<LONG>(sc << 16), keyname + length, 255 - length);
    };

    if (mods & MOD_SHIFT)   appendKey(VK_SHIFT);
    if (mods & MOD_CONTROL) appendKey(VK_CONTROL);
    if (mods & MOD_ALT)     appendKey(VK_MENU);
    if (mods & MOD_WIN)     appendKey(VK_LWIN);
    if (code != 0)          appendKey(code);

    return wstring(L" (") + keyname + L")";
}

// ---------------------------------------------------------------------------
// Hotkey registration
// ---------------------------------------------------------------------------

static void HandleHotkeyError(const wstring& label)
{
    WCHAR errorStr[512] = {};
    wstring msg = label + L"\n\n";
    FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, nullptr, GetLastError(),
                   MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), errorStr, 511, nullptr);
    msg += errorStr;
    MessageBoxW(nullptr, msg.c_str(), L"Hotkey Error", MB_OK);
}

static void RemoveHotkeys(HWND hWnd)
{
    int count = gPrefs.GetCount();
    for (int i = 0; i < count && i < cMaxDevices; ++i)
        if (gPrefs.GetHotkeyEnabled(i))
            UnregisterHotKey(hWnd, i);
    UnregisterHotKey(hWnd, CYCLE_HOTKEY);
}

static void InstallHotkeys(HWND hWnd)
{
    int count = gPrefs.GetCount();
    for (int i = 0; i < count && i < cMaxDevices; ++i)
    {
        if (!gPrefs.GetIsPresent(i) || !gPrefs.GetHotkeyEnabled(i))
            continue;
        UnregisterHotKey(hWnd, i);
        if (!RegisterHotKey(hWnd, i,
                            gPrefs.GetHotkeyMods(i) | gHotkeyFlags,
                            gPrefs.GetHotkeyCode(i)))
            HandleHotkeyError(gPrefs.GetName(i));
    }
    if (gPrefs.GetCycleKeyEnabled())
    {
        UnregisterHotKey(hWnd, CYCLE_HOTKEY);
        if (!RegisterHotKey(hWnd, CYCLE_HOTKEY,
                            gPrefs.GetCycleKeyMods() | gHotkeyFlags,
                            gPrefs.GetCycleKeyCode()))
            HandleHotkeyError(L"Cycle Hotkey");
    }
}

// ---------------------------------------------------------------------------
// Device cycling
// ---------------------------------------------------------------------------

static void CycleDevices()
{
    int count = gPrefs.GetCount();
    int start = gPrefs.FindByID(GetDefaultAudioPlaybackDevice()) + 1;

    for (int i = 0; i < count; ++i)
    {
        int j = (start + i) % count;
        if (!gPrefs.GetExcludeFromCycle(j) &&
            !gPrefs.GetIsHidden(j) &&
             gPrefs.GetIsPresent(j))
        {
            SetDefaultAudioPlaybackDevice(gPrefs.GetID(j).c_str());
            break;
        }
    }
}

// ---------------------------------------------------------------------------
// Single-instance guard
// ---------------------------------------------------------------------------

static bool AlreadyRunning()
{
    constexpr WCHAR kMutexName[] = L"{cf5daad3-840e-419c-8be3-3131897500ee}";

    // Create-or-open the mutex. Declared static so the handle lives for the
    // duration of the process — Windows releases it automatically on exit,
    // which is exactly the lifetime we want for a single-instance guard.
    static HANDLE s_hMutex = CreateMutexW(nullptr, TRUE, kMutexName);

    if (s_hMutex == nullptr)
        return false; // couldn't create mutex — let this instance run

    if (GetLastError() == ERROR_ALREADY_EXISTS)
    {
        // Another instance owns it. Close our handle and report conflict.
        CloseHandle(s_hMutex);
        s_hMutex = nullptr;
        return true;
    }

    // We created and own the mutex — we are the first instance.
    // Do NOT call ReleaseMutex: releasing signals waiters the resource is
    // free, which is not the same as the process being gone.
    return false;
}

// ---------------------------------------------------------------------------
// Version string
// ---------------------------------------------------------------------------

static wstring BuildVersionString(HINSTANCE hInstance)
{
    WCHAR appPathName[MAX_PATH]       = {};
    WCHAR appName[MAX_LOADSTRING + 1] = {};

    GetModuleFileNameW(hInstance, appPathName, MAX_PATH);
    DWORD verSize = GetFileVersionInfoSizeW(appPathName, nullptr);
    if (verSize == 0)
        return {};

    vector<char> buf(verSize);
    if (!GetFileVersionInfoW(appPathName, 0, verSize, buf.data()))
        return {};

    VS_FIXEDFILEINFO* pInfo   = nullptr;
    UINT              infoLen = 0;
    if (!VerQueryValueW(buf.data(), L"\\",
                        reinterpret_cast<void**>(&pInfo), &infoLen))
        return {};

    LoadStringW(hInstance, IDS_APP_TITLE, appName, MAX_LOADSTRING);

    wstring s(appName);
    s += L" ";
    s += to_wstring(HIWORD(pInfo->dwProductVersionMS));
    s += L".";
    s += to_wstring(LOWORD(pInfo->dwProductVersionMS));
    s += L".";
    s += to_wstring(HIWORD(pInfo->dwProductVersionLS));
    if (LOWORD(pInfo->dwProductVersionLS) != 0)
    {
        s += L".";
        s += to_wstring(LOWORD(pInfo->dwProductVersionLS));
    }
    return s;
}

// ---------------------------------------------------------------------------
// wWinMain
// ---------------------------------------------------------------------------

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                      _In_opt_ HINSTANCE /*hPrevInstance*/,
                      _In_ LPWSTR      /*lpCmdLine*/,
                      _In_ int         nCmdShow)
{
    if (AlreadyRunning())
    {
        MessageBoxW(nullptr,
                    L"An instance of Audio Endpoint Switcher is already running.",
                    L"Audio Endpoint Switcher", MB_OK);
        return 0;
    }

    LoadStringW(hInstance, IDS_APP_TITLE,  gszTitle,              MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_AES,        gszWindowClass,        MAX_LOADSTRING);
    LoadStringW(hInstance, IDS_APPTOOLTIP, gszApplicationToolTip, MAX_LOADSTRING);

    ghInstance = hInstance;
    MyRegisterClass(hInstance);

    HRESULT hrCom = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hrCom))
    {
        MessageBoxW(nullptr, L"Failed to initialize COM.",
                    L"Audio Endpoint Switcher", MB_OK | MB_ICONERROR);
        return 0;
    }

    if (!gPrefs.Load())
    {
        gPrefs.Save();
        EnableAutoStart();
    }
    EnumerateDevices();

    if (!InitInstance(hInstance, nCmdShow))
    {
        CoUninitialize();
        return 0;
    }

    UpdateTrayTooltip();

    HACCEL hAccelTable = LoadAcceleratorsW(hInstance, MAKEINTRESOURCEW(IDC_AES));

    MSG msg = {};
    while (GetMessageW(&msg, nullptr, 0, 0))
    {
        if (!TranslateAcceleratorW(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    CoUninitialize();
    return static_cast<int>(msg.wParam);
}

// ---------------------------------------------------------------------------
// Window class registration & creation
// ---------------------------------------------------------------------------

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex   = {};
    wcex.cbSize        = sizeof(WNDCLASSEXW);
    wcex.style         = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc   = WndProc;
    wcex.hInstance     = hInstance;
    wcex.hIcon         = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_ICON1));
    wcex.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wcex.lpszMenuName  = MAKEINTRESOURCEW(IDC_AES);
    wcex.lpszClassName = gszWindowClass;
    wcex.hIconSm       = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_SMALL));
    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int /*nCmdShow*/)
{
    hInst = hInstance;

    HWND hWnd = CreateWindowW(gszWindowClass, gszTitle, WS_OVERLAPPEDWINDOW,
                              CW_USEDEFAULT, 0, CW_USEDEFAULT, 0,
                              nullptr, nullptr, hInstance, nullptr);
    if (!hWnd) return FALSE;

    hMainIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_SMALL));

    nidApp                  = {};
    nidApp.cbSize           = sizeof(NOTIFYICONDATA);
    nidApp.hWnd             = hWnd;
    nidApp.uID              = IDI_SMALL;
    nidApp.uFlags           = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nidApp.uCallbackMessage = WM_USER_SHELLICON;
    wcsncpy_s(nidApp.szTip, 64, gszApplicationToolTip, _TRUNCATE);
    if (!GetDeviceIcon(&nidApp.hIcon))
        nidApp.hIcon = hMainIcon;
    Shell_NotifyIconW(NIM_ADD, &nidApp);

    return TRUE;
}

// ---------------------------------------------------------------------------
// Window procedure
// ---------------------------------------------------------------------------

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    static UINT s_uTaskbarRestart = 0;

    switch (message)
    {
    case WM_HOTKEY:
        if (wParam == CYCLE_HOTKEY)
            CycleDevices();
        else
            SetDefaultAudioPlaybackDevice(gPrefs.GetID(static_cast<int>(wParam)).c_str());
        return 0;

    case WM_CREATE:
        s_uTaskbarRestart = RegisterWindowMessageW(L"TaskbarCreated");
        pClient = make_unique<CMMNotificationClient>(hWnd);
        InstallNotificationCallback();
        InstallHotkeys(hWnd);
        return 0;

    case WM_USER_NOTIFICATION_DEFAULT:
        DestroyIcon(nidApp.hIcon);
        if (!GetDeviceIcon(&nidApp.hIcon))
            nidApp.hIcon = hMainIcon;
        UpdateTrayTooltip();
        return 0;

    case WM_USER_NOTIFICATION_ADDED:
    case WM_USER_NOTIFICATION_REMOVED:
    case WM_USER_NOTIFICATION_CHANGED:
        gbDevicesChanged = true;
        EnumerateDevices();
        return 0;

    case WM_USER_SHELLICON:
        switch (LOWORD(lParam))
        {
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        {
            if (IsWindow(gOpenWindow))
            {
                SetForegroundWindow(gOpenWindow);
                return 0;
            }
            if (gbDevicesChanged)
            {
                EnumerateDevices();
                gbDevicesChanged = false;
            }
            POINT pt = {};
            GetCursorPos(&pt);

            HMENU hPopMenu = CreatePopupMenu();
            InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION | MF_STRING,
                        ID_MENU_ABOUT, L"About...");
            InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR | MF_BYPOSITION, 0, nullptr);

            int count = gPrefs.GetCount();
            for (int i = 0; i < count; ++i)
            {
                if (gPrefs.GetIsHidden(i) || !gPrefs.GetIsPresent(i))
                    continue;
                wstring menuStr = gPrefs.GetName(i);
                if (gPrefs.GetHotkeyEnabled(i))
                    menuStr += MakeHotkeyString(gPrefs.GetHotkeyCode(i),
                                                gPrefs.GetHotkeyMods(i));
                UINT flags = MF_BYPOSITION | MF_STRING;
                if (IsDefaultAudioPlaybackDevice(gPrefs.GetID(i)))
                    flags |= MF_CHECKED;
                InsertMenuW(hPopMenu, 0xFFFFFFFF, flags,
                            ID_MENU_DEVICES + i, menuStr.c_str());
            }

            InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR | MF_BYPOSITION, 0, nullptr);
            InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION | MF_STRING,
                        ID_MENU_SETTINGS, L"Settings...");
            InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR | MF_BYPOSITION, 0, nullptr);
            InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION | MF_STRING,
                        ID_MENU_EXIT, L"Quit");

            SetForegroundWindow(hWnd);
            TrackPopupMenuEx(hPopMenu,
                             TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_BOTTOMALIGN,
                             pt.x, pt.y, hWnd, nullptr);
            DestroyMenu(hPopMenu);
            return 0;
        }
        }
        break;

    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        switch (wmId)
        {
        case ID_MENU_ABOUT:
            DialogBoxW(hInst, MAKEINTRESOURCEW(IDD_DIALOG_ABOUT), hWnd, About);
            return 0;

        case ID_MENU_EXIT:
            Shell_NotifyIconW(NIM_DELETE, &nidApp);
            DestroyWindow(hWnd);
            return 0;

        case ID_MENU_SETTINGS:
            RemoveHotkeys(hWnd);
            DoSettingsDialog(hInst, hWnd);
            EnumerateDevices();
            InstallHotkeys(hWnd);
            return 0;

        default:
            if (wmId >= static_cast<int>(ID_MENU_DEVICES) &&
                wmId <  static_cast<int>(ID_MENU_DEVICES + cMaxDevices))
            {
                int devIdx = wmId - static_cast<int>(ID_MENU_DEVICES);
                SetDefaultAudioPlaybackDevice(gPrefs.GetID(devIdx).c_str());
                return 0;
            }
            return DefWindowProcW(hWnd, message, wParam, lParam);
        }
    }

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;

    case WM_POWERBROADCAST:
        if (wParam == PBT_APMRESUMEAUTOMATIC)
            SetTimer(hWnd, TIMER_AUDIO_REFRESH, 1000, nullptr);
        return TRUE;

    case WM_TIMER:
        if (wParam == TIMER_AUDIO_REFRESH)
        {
            KillTimer(hWnd, TIMER_AUDIO_REFRESH);
            EnumerateDevices();
            UpdateTrayTooltip();
        }
        return 0;

    default:
        if (s_uTaskbarRestart && message == s_uTaskbarRestart)
            Shell_NotifyIconW(NIM_ADD, &nidApp);
        return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return DefWindowProcW(hWnd, message, wParam, lParam);
}

// ---------------------------------------------------------------------------
// Settings dialog wrapper
// ---------------------------------------------------------------------------

void DoSettingsDialog(HINSTANCE hInst, HWND hWnd)
{
    CQSESPrefs temp = gPrefs;
    INT_PTR rc = DialogBoxParamW(hInst, MAKEINTRESOURCEW(IDD_DIALOG_SETTINGS),
                                  hWnd, SettingsDialogProc,
                                  reinterpret_cast<LPARAM>(&temp));
    if (rc == IDOK)
    {
        gPrefs = temp;
        gPrefs.Save();
    }
}

// ---------------------------------------------------------------------------
// About dialog
// ---------------------------------------------------------------------------

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM /*lParam*/)
{
    switch (message)
    {
    case WM_INITDIALOG:
        gOpenWindow = hDlg;
        SetWindowTextW(GetDlgItem(hDlg, IDC_STATIC_APP),
                       BuildVersionString(ghInstance).c_str());
        return TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return TRUE;
        }
        break;

    case WM_DESTROY:
        gOpenWindow = nullptr;
        return TRUE;
    }
    return FALSE;
}

// ---------------------------------------------------------------------------
// Device icon extraction
// ---------------------------------------------------------------------------

bool GetDeviceIcon(HICON* hIcon)
{
    if (!hIcon) return false;
    *hIcon = nullptr;

    CComPtr<IMMDeviceEnumerator> pEnum;
    CComPtr<IMMDevice>           pDevice;
    CComPtr<IPropertyStore>      pStore;
    PROPVARIANT pv;
    PropVariantInit(&pv);

    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                                __uuidof(IMMDeviceEnumerator),
                                reinterpret_cast<void**>(&pEnum))))
        return false;

    if (FAILED(pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice)))
        return false;

    if (FAILED(pDevice->OpenPropertyStore(STGM_READ, &pStore)))
        return false;

    if (FAILED(pStore->GetValue(PKEY_DeviceClass_IconPath, &pv)))
        return false;

    wstring iconPath(pv.pwszVal);
    PropVariantClear(&pv);

    size_t commaPos = iconPath.find(L',');
    if (commaPos == wstring::npos)
        return false;

    wstring exePath = iconPath.substr(0, commaPos);
    int     nIconID = _wtoi(iconPath.c_str() + commaPos + 1);

    HICON hLarge = nullptr;
    ExtractIconExW(exePath.c_str(), nIconID, &hLarge, nullptr, 1);
    if (!hLarge)
        return false;

    int cx = GetSystemMetrics(SM_CXSMICON);
    int cy = GetSystemMetrics(SM_CYSMICON);

    HDC hdc    = GetDC(nullptr);
    HDC hdcMem = CreateCompatibleDC(hdc);

    BITMAPINFO bmi              = {};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = cx;
    bmi.bmiHeader.biHeight      = -cy;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void*   pvBits = nullptr;
    HBITMAP hDIB   = CreateDIBSection(hdc, &bmi, DIB_RGB_COLORS, &pvBits, nullptr, 0);
    if (hDIB)
    {
        HBITMAP hOld = static_cast<HBITMAP>(SelectObject(hdcMem, hDIB));
        DrawIconEx(hdcMem, 0, 0, hLarge, cx, cy, 0, nullptr, DI_NORMAL);

        ICONINFO ii = {};
        ii.fIcon    = TRUE;
        ii.hbmColor = hDIB;
        ii.hbmMask  = CreateBitmap(cx, cy, 1, 1, nullptr);
        *hIcon      = CreateIconIndirect(&ii);
        DeleteObject(ii.hbmMask);

        SelectObject(hdcMem, hOld);
        DeleteObject(hDIB);
    }

    DeleteDC(hdcMem);
    ReleaseDC(nullptr, hdc);
    DestroyIcon(hLarge);

    return (*hIcon != nullptr);
}

// ---------------------------------------------------------------------------
// Audio endpoint helpers
// ---------------------------------------------------------------------------

bool SetDefaultAudioPlaybackDevice(LPCWSTR devID)
{
    CComPtr<IPolicyConfig> pPolicyConfig;
    HRESULT hr = pPolicyConfig.CoCreateInstance(__uuidof(CPolicyConfigClient));
    if (SUCCEEDED(hr))
        hr = pPolicyConfig->SetDefaultEndpoint(devID, eConsole);
    return SUCCEEDED(hr);
}

void UpdateTrayTooltip()
{
    wstring id   = GetDefaultAudioPlaybackDevice();
    int     idx  = gPrefs.FindByID(id);
    wstring name = (idx >= 0) ? gPrefs.GetName(idx) : wstring(gszApplicationToolTip);
    wcsncpy_s(nidApp.szTip, 64, name.c_str(), _TRUNCATE);
    Shell_NotifyIconW(NIM_MODIFY, &nidApp);
}

void InstallNotificationCallback()
{
    CComPtr<IMMDeviceEnumerator> pEnum;
    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                                  __uuidof(IMMDeviceEnumerator),
                                  reinterpret_cast<void**>(&pEnum));
    if (SUCCEEDED(hr))
        pEnum->RegisterEndpointNotificationCallback(pClient.get());
}

wstring GetDefaultAudioPlaybackDevice()
{
    CComPtr<IMMDeviceEnumerator> pEnum;
    CComPtr<IMMDevice>           pDevice;
    LPWSTR pwszID = nullptr;

    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                                __uuidof(IMMDeviceEnumerator),
                                reinterpret_cast<void**>(&pEnum))))
        return {};

    if (FAILED(pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice)))
        return {};

    if (FAILED(pDevice->GetId(&pwszID)))
        return {};

    wstring id(pwszID);
    CoTaskMemFree(pwszID);
    return id;
}

bool IsDefaultAudioPlaybackDevice(const wstring& s)
{
    return GetDefaultAudioPlaybackDevice() == s;
}

// ---------------------------------------------------------------------------
// Device enumeration
// ---------------------------------------------------------------------------

int EnumerateDevices()
{
    CComPtr<IMMDeviceEnumerator> pEnum;
    CComPtr<IMMDeviceCollection> pDevices;

    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL,
                                  __uuidof(IMMDeviceEnumerator),
                                  reinterpret_cast<void**>(&pEnum));
    if (FAILED(hr)) return -1;

    hr = pEnum->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);
    if (FAILED(hr)) return -1;

    UINT count = 0;
    pDevices->GetCount(&count);
    if (count > static_cast<UINT>(cMaxDevices))
        count = static_cast<UINT>(cMaxDevices);

    gPrefs.ResetPresent();

    for (UINT i = 0; i < count; ++i)
    {
        CComPtr<IMMDevice>      pDevice;
        CComPtr<IPropertyStore> pStore;
        PROPVARIANT pv;
        PropVariantInit(&pv);

        if (FAILED(pDevices->Item(i, &pDevice))) continue;

        LPWSTR wstrID = nullptr;
        if (FAILED(pDevice->GetId(&wstrID))) continue;

        DevicePrefs dev;
        dev.DeviceID = wstrID;
        CoTaskMemFree(wstrID);

        if (FAILED(pDevice->OpenPropertyStore(STGM_READ, &pStore))) continue;
        if (FAILED(pStore->GetValue(PKEY_Device_FriendlyName, &pv)))
        {
            PropVariantClear(&pv);
            continue;
        }
        dev.Name = pv.pwszVal;
        PropVariantClear(&pv);

        gPrefs.Update(dev);
    }

    gPrefs.Sort();
    return 0;
}
