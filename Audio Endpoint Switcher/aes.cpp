#include <wchar.h>
#include "windows.h"
#include <VersionHelpers.h>
//#include <shellapi.h>
#include <string>
#include <sstream>
#include <Mmdeviceapi.h>
#include <Shlobj.h>
#include <atlbase.h>
#include <atlcomcli.h>   // defines CComPtr

#include "PolicyConfig.h"
#include "Propidl.h"
#include "Functiondiscoverykeys_devpkey.h"
#include "aesIMMNotificationClient.h"

#include "Prefs.h"
#include "Settings.h"
#include "resource.h"

using namespace std;

#define MAX_LOADSTRING 100
#define	WM_USER_SHELLICON WM_USER + 1

#define ID_MENU_ABOUT (UINT)100
#define ID_MENU_DONATE (UINT)101
#define ID_MENU_DEVICES (UINT)120
#define ID_MENU_EXIT (UINT)140
#define ID_MENU_SETTINGS (UINT)150
#define CYCLE_HOTKEY (UINT)200
#define TIMER_AUDIO_REFRESH (UINT)1234

// Global Variables:
HINSTANCE hInst;
NOTIFYICONDATA nidApp;
HICON hMainIcon;
bool gbDevicesChanged = false;

WCHAR gszTitle[MAX_LOADSTRING];
WCHAR gszWindowClass[MAX_LOADSTRING];
WCHAR gszApplicationToolTip[MAX_LOADSTRING];
HINSTANCE ghInstance;
HWND gOpenWindow = 0;
const UINT gHotkeyFlags = MOD_NOREPEAT;
wstringstream gVersionString;

CQSESPrefs gPrefs;
CMMNotificationClient * pClient = 0;


// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	HotkeyDlg(HWND, UINT, WPARAM, LPARAM);
void DoSettingsDialog(HINSTANCE hInst, HWND hWnd);

void UpdateTrayTooltip();
void InstallNotificationCallBack();
int EnumerateDevices();
bool SetDefaultAudioPlaybackDevice(LPCWSTR devID);
bool GetDeviceIcon(HICON * hIcon);
bool IsDefaultAudioPlaybackDevice(const wstring& s);
wstring& GetDefaultAudioPlaybackDevice(wstring& s);

void MakeStartupLinkPath(WCHAR * name, WCHAR * path)
{
	LPWSTR wstrStartupPath = 0;
	HRESULT hres = SHGetKnownFolderPath(FOLDERID_Startup, 0, 0, &wstrStartupPath);
	if (SUCCEEDED(hres))
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
	IShellLink* pShellLink = NULL;
	IPersistFile* pPersistFile;
	HRESULT hres;

	hres = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_ALL, IID_IShellLink, (void**)&pShellLink);
	if (SUCCEEDED(hres))
	{
		hres = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);
		if (SUCCEEDED(hres))
		{
			HMODULE hModule = GetModuleHandle(NULL);
			WCHAR szExeName[MAX_PATH];
			GetModuleFileNameW(hModule, szExeName, MAX_PATH);
			pShellLink->SetPath(szExeName);
			pShellLink->SetDescription(gszApplicationToolTip);
			pShellLink->SetIconLocation(szExeName, 0);
			hres = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);
			if (SUCCEEDED(hres))
			{
				WCHAR szStartupLinkName[MAX_PATH];
				MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
				pPersistFile->Save(szStartupLinkName, TRUE);
				pPersistFile->Release();
			}
			pShellLink->Release();
		}
	}
}

void DisableAutoStart(void)
{
	WCHAR szStartupLinkName[MAX_PATH];

	MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
	if (*szStartupLinkName != L'\0')
		DeleteFileW(szStartupLinkName);
}

UINT CheckAutoStart(void)
{
	WCHAR szStartupLinkName[MAX_PATH];

	MakeStartupLinkPath(gszApplicationToolTip, szStartupLinkName);
	HANDLE hFile = CreateFileW(szStartupLinkName, GENERIC_READ, 0, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (INVALID_HANDLE_VALUE != hFile)
	{
		CloseHandle(hFile);
		return 1;
	}
	return 0;
}

wstring& MakeHotkeyString(UINT code, UINT mods, wstring & str)
{
	WCHAR keyname[256] = {};
	UINT scanCode;
	int length = 0;

	if (mods & MOD_SHIFT)
	{
		scanCode = MapVirtualKey(VK_SHIFT, MAPVK_VK_TO_VSC);
		length = GetKeyNameTextW(scanCode << 16, keyname, 255);
	}
	if (mods & MOD_CONTROL)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(VK_CONTROL, MAPVK_VK_TO_VSC);
		length += GetKeyNameTextW(scanCode << 16, keyname + length, 255 - length);
	}
	if (mods & MOD_ALT)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(VK_MENU, MAPVK_VK_TO_VSC);
		length += GetKeyNameTextW(scanCode << 16, keyname + length, 255 - length);
	}
	if (mods & MOD_WIN)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(VK_LWIN, MAPVK_VK_TO_VSC);
		length += GetKeyNameTextW(scanCode << 16, keyname + length, 255 - length);
	}
	if (code != 0)
	{
		if (length > 0) { keyname[length++] = L'+'; keyname[length] = 0; }
		scanCode = MapVirtualKey(code, MAPVK_VK_TO_VSC);
		length += GetKeyNameTextW(scanCode << 16, keyname + length, 255 - length);
	}
	str += L" (";
	str += keyname;
	return str += L")";
}

void RemoveHotkeys(HWND hWnd)
{
	UINT count = gPrefs.GetCount();

	for (UINT i = 0; i < count  && i < cMaxDevices; i++)
	{
		if (gPrefs.GetHotkeyEnabled(i))
		{
			UnregisterHotKey(hWnd,i);
		}
	}
	UnregisterHotKey(hWnd, CYCLE_HOTKEY);
}

// Message Box with human readable error if hotkey fails to register.
void HandleHotkeyError(wstring& s)
{
	WCHAR ErrorStr[512] = {0};
	s += L"\n\n";
	DWORD err = GetLastError();
	FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, 0, err, MAKELANGID(LANG_NEUTRAL,SUBLANG_DEFAULT), ErrorStr, 511, 0);
	s += ErrorStr;
	MessageBoxW(NULL, s.c_str(), L"Hotkey Error", MB_OK);
}

// Install all hotkeys. Do this on load or when the prefs change.
void InstallHotkeys(HWND hWnd)
{
	UINT count = gPrefs.GetCount();
	wstring s;

	for (UINT i = 0; i < count  && i < cMaxDevices; i++)
	{
		if (!gPrefs.GetIsPresent(i))
			continue;
		if (gPrefs.GetHotkeyEnabled(i))
		{
			UnregisterHotKey(hWnd,i);
			if(0 == RegisterHotKey(hWnd, i, gPrefs.GetHotkeyMods(i) | gHotkeyFlags, gPrefs.GetHotkeyCode(i)))
				HandleHotkeyError(gPrefs.GetName(i, s));
		}
	}
	if (gPrefs.GetCycleKeyEnabled())
	{
		UnregisterHotKey(hWnd, CYCLE_HOTKEY);
		if (0 == RegisterHotKey(hWnd, CYCLE_HOTKEY, gPrefs.GetCycleKeyMods() | gHotkeyFlags, gPrefs.GetCycleKeyCode()))
		{
			s = L"Cycle Hotkey";
			HandleHotkeyError(s);
		}
	}
}

// Function to cycle through visible and cycle-enabled devices.
void CycleDevices()
{
	wstring IdString;
	UINT count = gPrefs.GetCount();
	UINT start = gPrefs.FindByID(GetDefaultAudioPlaybackDevice(IdString)) + 1;

	for (UINT i = 0, j; i < count; i++)
	{
		j = start++ % count;
		gPrefs.GetName(j, IdString);
		if(!gPrefs.GetExcludeFromCycle(j) && !gPrefs.GetIsHidden(j) && gPrefs.GetIsPresent(j))
		{
			SetDefaultAudioPlaybackDevice((gPrefs.GetID(j, IdString)).c_str());
			break;
		}
	}
}

// Check to see if we're already running.
bool AlreadyRunning()
{
#define APPLICATION_INSTANCE_MUTEX_NAME L"{cf5daad3-840e-419c-8be3-3131897500ee}"

    //Make sure at most one instance of the tool is running
    HANDLE hMutexOneInstance(::CreateMutex(NULL, TRUE, APPLICATION_INSTANCE_MUTEX_NAME));
    bool bAlreadyRunning((::GetLastError() == ERROR_ALREADY_EXISTS));
    if (hMutexOneInstance == NULL || bAlreadyRunning)
    {
    	if(hMutexOneInstance)
    	{
    		::ReleaseMutex(hMutexOneInstance);
    		::CloseHandle(hMutexOneInstance);
    	}
    }
	return bAlreadyRunning;
}

// Check OS version - requires Windows 7 or later.
bool CheckOSVersion()
{
	if (!IsWindows7OrGreater())
	{
		MessageBoxW(NULL, L"Requires Windows 7 or later.", L"Audio Endpoint Switcher", MB_OK);
		return false;
	}
	return true;
}

void MakegVersionString(HINSTANCE hInstance)
{
	VS_FIXEDFILEINFO *pFixedInfo;	// pointer to fixed file info structure
	UINT    uVersionLen;			// Current length of full version string
	WCHAR AppPathName[MAX_PATH];
	WCHAR AppName[MAX_LOADSTRING+1];

	ZeroMemory(AppName, (MAX_LOADSTRING+1) * sizeof(WCHAR));

	GetModuleFileNameW(hInstance, AppPathName, MAX_PATH);
	DWORD dwVerInfoSize = GetFileVersionInfoSizeW(AppPathName, NULL);
	void *pBuffer = new char [dwVerInfoSize];
	if(pBuffer == 0)
		return;
	GetFileVersionInfoW(AppPathName, 0, dwVerInfoSize, pBuffer);
	VerQueryValueW(pBuffer, L"\\", (void**)&pFixedInfo, (UINT *)&uVersionLen);

	LoadStringW(hInstance, IDS_APP_TITLE, AppName, MAX_LOADSTRING);
	gVersionString.str(L"");
	gVersionString.clear();
	gVersionString << AppName
		<< L" " << HIWORD (pFixedInfo->dwProductVersionMS)
		<< L"." << LOWORD (pFixedInfo->dwProductVersionMS)
		<< L"." << HIWORD (pFixedInfo->dwProductVersionLS);
	if(LOWORD (pFixedInfo->dwProductVersionLS) != 0)
		gVersionString << L"." << LOWORD (pFixedInfo->dwProductVersionLS);
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
					  _In_opt_ HINSTANCE hPrevInstance,
					  _In_ LPWSTR     lpCmdLine,
					  _In_ int        nCmdShow)
{
	if (!CheckOSVersion())
		return 0;

	if (AlreadyRunning())
	{
		MessageBoxW(NULL, L"An instance of Audio Endpoint Switcher is already running.", L"Audio Endpoint Switcher", MB_OK);
		return false;
	}

	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadStringW(hInstance, IDS_APP_TITLE, gszTitle, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_AES, gszWindowClass, MAX_LOADSTRING);
	LoadStringW(hInstance, IDS_APPTOOLTIP, gszApplicationToolTip, MAX_LOADSTRING);

	ghInstance = hInstance;
	MyRegisterClass(hInstance);

	HRESULT hrCom = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hrCom))
	{
		MessageBoxW(NULL, L"Failed to initialize COM.", L"Audio Endpoint Switcher", MB_OK | MB_ICONERROR);
		return FALSE;
	}

	gPrefs.Init(10);
	if(!gPrefs.Load())
	{
		gPrefs.Save();
		EnableAutoStart();
	}
	EnumerateDevices();

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	// Update tooltip to show current default device name now that devices are enumerated.
	UpdateTrayTooltip();

	hAccelTable = LoadAcceleratorsW(hInstance, MAKEINTRESOURCEW(IDC_AES));

	// Main message loop:
	while (int ret = GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	CoUninitialize();
	return (int) msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEXW wcex;

	wcex.cbSize = sizeof(WNDCLASSEXW);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_ICON1));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCEW(IDC_AES);
	wcex.lpszClassName	= gszWindowClass;
	wcex.hIconSm		= LoadIconW(wcex.hInstance, MAKEINTRESOURCEW(IDI_SMALL));

	return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND hWnd;

	hInst = hInstance; // Store instance handle in our global variable

	hWnd = CreateWindowW(gszWindowClass, gszTitle, WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

	if (!hWnd)
		return FALSE;

	hMainIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_SMALL));

	nidApp.cbSize = sizeof(NOTIFYICONDATA);
	nidApp.hWnd = (HWND) hWnd;
	nidApp.uID = IDI_SMALL; 
	nidApp.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	nidApp.uCallbackMessage = WM_USER_SHELLICON;
	wcsncpy_s(nidApp.szTip, 64, gszApplicationToolTip, _TRUNCATE); // initial value, updated after EnumerateDevices
	if (!GetDeviceIcon(&nidApp.hIcon))
		nidApp.hIcon = hMainIcon;
	Shell_NotifyIconW(NIM_ADD, &nidApp); 

	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	POINT lpClickPoint;
	UINT count;
	HMENU hPopMenu;
	wstring MenuString, TempString;

	static UINT s_uTaskbarRestart;

	switch (message)
	{
	case WM_HOTKEY:
		if (wParam == CYCLE_HOTKEY)
			CycleDevices();
		else
			SetDefaultAudioPlaybackDevice((gPrefs.GetID((int)wParam, TempString)).c_str());
		return TRUE;
	case WM_CREATE:
		pClient = new CMMNotificationClient(hWnd);
		// load the preferences and fill in audio devices
		InstallNotificationCallBack();
		InstallHotkeys(hWnd);
		return TRUE;
	case WM_USER_NOTIFICATION_DEFAULT:
		DestroyIcon(nidApp.hIcon);
		if (!GetDeviceIcon(&nidApp.hIcon))
			nidApp.hIcon = hMainIcon;
		UpdateTrayTooltip();
		return TRUE;
	case WM_USER_NOTIFICATION_ADDED:
	case WM_USER_NOTIFICATION_REMOVED:
	case WM_USER_NOTIFICATION_CHANGED:
		gbDevicesChanged = true;
		EnumerateDevices();
		return TRUE;
	case WM_USER_SHELLICON: 
		// systray msg callback 
		switch(LOWORD(lParam)) 
		{   
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
			if (IsWindow(gOpenWindow))
			{
				SetForegroundWindow(gOpenWindow);
				return TRUE;
			}
			if (gbDevicesChanged)
			{
				EnumerateDevices();
				gbDevicesChanged = false;
			}
			GetCursorPos(&lpClickPoint);
			hPopMenu = CreatePopupMenu();
			InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_ABOUT, L"About...");
			InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR|MF_BYPOSITION, 0, NULL);

			count = gPrefs.GetCount();
			for (UINT i = 0; i < count; i++)
			{
				if (gPrefs.GetIsHidden(i) || !gPrefs.GetIsPresent(i))
					continue;
				gPrefs.GetName(i, MenuString);
				if (gPrefs.GetHotkeyEnabled(i))
					MenuString = MakeHotkeyString(gPrefs.GetHotkeyCode(i), gPrefs.GetHotkeyMods(i), MenuString);
				if (IsDefaultAudioPlaybackDevice(gPrefs.GetID(i, TempString)))
					InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION|MF_STRING|MF_CHECKED, ID_MENU_DEVICES + i, MenuString.c_str());
				else
					InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_DEVICES + i, MenuString.c_str());
			}

			gPrefs.GetCycleKeyString(TempString);

			InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR|MF_BYPOSITION, 0, NULL);
			InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_SETTINGS, L"Settings...");
			InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_SEPARATOR|MF_BYPOSITION, 0, NULL);
			InsertMenuW(hPopMenu, 0xFFFFFFFF, MF_BYPOSITION|MF_STRING, ID_MENU_EXIT, L"Quit");
			//								
			SetForegroundWindow(hWnd);
			TrackPopupMenu(hPopMenu,TPM_LEFTALIGN|TPM_LEFTBUTTON|TPM_BOTTOMALIGN,lpClickPoint.x, lpClickPoint.y,0,hWnd,NULL);
			return TRUE; 
		}
		break;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);

		switch (wmId)
		{
		case ID_MENU_ABOUT:
			DialogBoxW(hInst, MAKEINTRESOURCEW(IDD_DIALOG_ABOUT), hWnd, About);
			return TRUE;
		case ID_MENU_EXIT:
			Shell_NotifyIconW(NIM_DELETE, &nidApp);
			SendMessage(hWnd, WM_CLOSE, 0, 0);
			DestroyWindow(hWnd);
			return TRUE;
		case ID_MENU_SETTINGS:
			RemoveHotkeys(hWnd);
			DoSettingsDialog(hInst, hWnd);
			EnumerateDevices();
			InstallHotkeys(hWnd);
			return TRUE;
		default:
			if (wmId >= (ID_MENU_DEVICES) && wmId < (ID_MENU_DEVICES + cMaxDevices))
			{
				UINT devID = wmId - ID_MENU_DEVICES;
				SetDefaultAudioPlaybackDevice((gPrefs.GetID(devID, TempString)).c_str());
				return TRUE;
			}
			else
				return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;

	case WM_POWERBROADCAST:
		if (wParam == PBT_APMRESUMEAUTOMATIC) {
			SetTimer(hWnd, TIMER_AUDIO_REFRESH, 1000, nullptr); // delay 1s
		}
		break;

	case WM_TIMER:
		if (wParam == TIMER_AUDIO_REFRESH) {
			KillTimer(hWnd, TIMER_AUDIO_REFRESH);
			EnumerateDevices(); // refresh device list
			UpdateTrayTooltip();
		}
		break;

	default:
		if(message == s_uTaskbarRestart)
			Shell_NotifyIconW(NIM_ADD, &nidApp);
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return FALSE;
}

void DoSettingsDialog(HINSTANCE hInst, HWND hWnd)
{
	CQSESPrefs TempPrefs;
	INT_PTR rc;

	TempPrefs = gPrefs;
	rc = DialogBoxParamW(hInst, MAKEINTRESOURCEW(IDD_DIALOG_SETTINGS), hWnd, SettingsDialogProc, (LPARAM) &TempPrefs);

	if (rc == IDOK)
	{
		gPrefs = TempPrefs;
		gPrefs.Save();
	}
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		gOpenWindow = hDlg;
		MakegVersionString(ghInstance);
		SetWindowTextW(GetDlgItem(hDlg, IDC_STATIC_APP), gVersionString.str().c_str());
		return (INT_PTR)TRUE;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	case WM_DESTROY:
		gOpenWindow = 0;
		return (INT_PTR)TRUE;
	}
	return (INT_PTR)FALSE;
}

#define EXIT_ON_ERROR(hres)  \
              if (FAILED(hres)) { goto Exit; }

bool GetDeviceIcon(HICON* hIcon)
{
	if (!hIcon)
		return false;

	*hIcon = NULL;

	wstring path;
	CComPtr<IMMDeviceEnumerator> pEnum;
	CComPtr<IPropertyStore> pStore;
	CComPtr<IMMDevice> pDevice;
	PROPVARIANT propVariant;
	PropVariantInit(&propVariant);
	HRESULT hr;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	if (FAILED(hr))
		return false;

	hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
	if (FAILED(hr))
		return false;

	hr = pDevice->OpenPropertyStore(STGM_READ, &pStore);
	if (FAILED(hr))
		return false;

	hr = pStore->GetValue(PKEY_DeviceClass_IconPath, &propVariant);
	if (FAILED(hr))
		return false;

	path = propVariant.pwszVal;
	PropVariantClear(&propVariant);

	wstring wsIconPath = path;
	size_t pos = wsIconPath.find(L',');
	if (pos != std::string::npos)
	{
		wstring wsIconId = wsIconPath.substr(pos + 1, string::npos);
		wsIconPath.erase(pos, string::npos);
		int nIconID = _wtoi(wsIconId.c_str());

		// Extract the large icon and scale it down to the systray size for best quality.  
		// Requesting the small icon directly gives a low-res 16x16 that looks poor on
		// modern displays. Extracting large and scaling produces a much sharper result.
		HICON hLargeIcon = NULL;
		if (nIconID < 0)
			ExtractIconExW(wsIconPath.c_str(), nIconID, &hLargeIcon, NULL, 1);
		else
			ExtractIconExW(wsIconPath.c_str(), nIconID, &hLargeIcon, NULL, 1);

		if (hLargeIcon)
		{
			// Scale the large icon down to the system-defined small icon size.
			int cx = GetSystemMetrics(SM_CXSMICON);
			int cy = GetSystemMetrics(SM_CYSMICON);
			HDC hdc = GetDC(NULL);
			HDC hdcMem = CreateCompatibleDC(hdc);
			HBITMAP hBmp = CreateCompatibleBitmap(hdc, cx, cy);
			HBITMAP hOldBmp = (HBITMAP)SelectObject(hdcMem, hBmp);

			BITMAPINFO bmi = {};
			bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
			bmi.bmiHeader.biWidth = cx;
			bmi.bmiHeader.biHeight = -cy;
			bmi.bmiHeader.biPlanes = 1;
			bmi.bmiHeader.biBitCount = 32;
			bmi.bmiHeader.biCompression = BI_RGB;
			void* pvBits = NULL;
			HBITMAP hDIB = CreateDIBSection(hdc, &bmi, DIB_RGB_COLORS, &pvBits, NULL, 0);
			if (hDIB)
			{
				HBITMAP hOldDIB = (HBITMAP)SelectObject(hdcMem, hDIB);

				DrawIconEx(hdcMem, 0, 0, hLargeIcon, cx, cy, 0, NULL, DI_NORMAL);

				ICONINFO ii = {};
				ii.fIcon = TRUE;
				ii.hbmColor = hDIB;
				ii.hbmMask = CreateBitmap(cx, cy, 1, 1, NULL);
				*hIcon = CreateIconIndirect(&ii);

				DeleteObject(ii.hbmMask);
				SelectObject(hdcMem, hOldDIB);
				DeleteObject(hDIB);
			}
			SelectObject(hdcMem, hOldBmp);
			DeleteObject(hBmp);
			DeleteDC(hdcMem);
			ReleaseDC(NULL, hdc);
			DestroyIcon(hLargeIcon);
		}
	}

	return (*hIcon != NULL && *hIcon != (HICON)1);
}

bool SetDefaultAudioPlaybackDevice(LPCWSTR devID)
{	
	CComPtr<IPolicyConfig> pPolicyConfig;
	ERole reserved = eConsole;
	HRESULT hr;

	hr = pPolicyConfig.CoCreateInstance(__uuidof(CPolicyConfigClient));
	if (SUCCEEDED(hr))
	{
		hr = pPolicyConfig->SetDefaultEndpoint(devID, reserved);
	}
	return SUCCEEDED(hr);
}

void UpdateTrayTooltip()
{
	wstring id, name;
	int index = gPrefs.FindByID(GetDefaultAudioPlaybackDevice(id));
	if (index >= 0)
		gPrefs.GetName(index, name);
	else
		name = gszApplicationToolTip; // fallback to app name if device not found
	wcsncpy_s(nidApp.szTip, 64, name.c_str(), _TRUNCATE);
	Shell_NotifyIconW(NIM_MODIFY, &nidApp);
}

void InstallNotificationCallBack()
{
	HRESULT hr;

	IMMDeviceEnumerator *pEnum = NULL;
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	if (SUCCEEDED(hr))
	{
		pEnum->RegisterEndpointNotificationCallback(pClient);
		pEnum->Release();
	}
}

wstring& GetDefaultAudioPlaybackDevice(wstring& s)
{
	HRESULT hr;
	CComPtr<IMMDeviceEnumerator> pEnum;
	CComPtr<IMMDevice> pDefDevice;
	LPWSTR pwszDefID = NULL;

	s.clear();
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	if (FAILED(hr))
		return s;

	hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDefDevice);
	if (FAILED(hr))
		return s;

	hr = pDefDevice->GetId(&pwszDefID);
	if (FAILED(hr))
		return s;

	s = pwszDefID;
	CoTaskMemFree(pwszDefID);
	return s;
}

bool IsDefaultAudioPlaybackDevice(const wstring& s)
{
	HRESULT hr;
	CComPtr<IMMDeviceEnumerator> pEnum;
	CComPtr<IMMDevice> pDefDevice;
	LPWSTR pwszDefID = NULL;

	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	if (FAILED(hr))
		return false;

	hr = pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDefDevice);
	if (FAILED(hr))
		return false;

	hr = pDefDevice->GetId(&pwszDefID);
	if (FAILED(hr))
		return false;

	bool bMatch = (s == pwszDefID);
	CoTaskMemFree(pwszDefID);
	return bMatch;
}

int EnumerateDevices()
{
	LPWSTR wstrID = NULL;
	UINT count = 0;
	DevicePrefs SoundDeviceInfo;

	wstring wsDefaultDeviceID(L"");

	SoundDeviceInfo.DeviceID = L"";
	SoundDeviceInfo.HotkeyString = L"";
	SoundDeviceInfo.HasHotkey = false;
	SoundDeviceInfo.IsExcludedFromCycle = false;
	SoundDeviceInfo.IsHidden = false;
	SoundDeviceInfo.IsPresent = false;
	SoundDeviceInfo.KeyCode = 0;
	SoundDeviceInfo.KeyMods = 0;

	HRESULT hr;

	// Use smart COM pointers for RAII and safer Release semantics.
	CComPtr<IMMDeviceEnumerator> pEnum;
	CComPtr<IMMDeviceCollection> pDevices;
	CComPtr<IMMDevice> pDevice;
	CComPtr<IPropertyStore> pStore;
	PROPVARIANT propVariant;
	PropVariantInit(&propVariant);

	// Create a multimedia device enumerator.
	hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
		CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum);
	EXIT_ON_ERROR(hr);

	// Enumerate the output devices.
	hr = pEnum->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pDevices);
	EXIT_ON_ERROR(hr);

	hr = pDevices->GetCount(&count);
	EXIT_ON_ERROR(hr);

	if (count > cMaxDevices)
		count = cMaxDevices;

	gPrefs.ResetPresent(); // mark all devices not present

	for (UINT i = 0; i < count; i++)
	{
		// Ensure local variables are clean for each iteration.
		wstrID = NULL;
		pStore.Release();
		pDevice.Release();

		hr = pDevices->Item(i, &pDevice);
		EXIT_ON_ERROR(hr);

		hr = pDevice->GetId(&wstrID);
		EXIT_ON_ERROR(hr);
		SoundDeviceInfo.DeviceID = wstrID;
		CoTaskMemFree(wstrID);
		wstrID = NULL;

		hr = pDevice->OpenPropertyStore(STGM_READ, &pStore);
		EXIT_ON_ERROR(hr);

		PropVariantInit(&propVariant);
		hr = pStore->GetValue(PKEY_Device_FriendlyName, &propVariant);
		if ((((HRESULT)(hr)) < 0)) {
			PropVariantClear(&propVariant);
			goto Exit;
		};

		SoundDeviceInfo.Name = propVariant.pwszVal;
		gPrefs.Update(&SoundDeviceInfo);

		// release per-iteration references so Exit cleanup only deals with remaining objects
		PropVariantClear(&propVariant);
		pStore.Release();
		pDevice.Release();
	}
Exit:
	// make sure any leftover allocations are freed
	CoTaskMemFree(wstrID);

	// smart pointers auto-release on scope exit; explicit release here for clarity
	pStore.Release();
	pDevice.Release();
	pDevices.Release();
	pEnum.Release();

	gPrefs.Sort();
	return 0;
}
