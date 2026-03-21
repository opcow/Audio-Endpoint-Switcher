#include <Windowsx.h>
#include <tchar.h>
#include <string>
#include <wchar.h>
#include "resource.h"
#include "Settings.h"

using namespace std;

typedef LRESULT (__stdcall * CONTROLPROC) (HWND, UINT, WPARAM, LPARAM);

extern HWND gOpenWindow;
extern void EnableAutoStart();
extern void DisableAutoStart(void);
extern UINT CheckAutoStart(void);

static CQSESPrefs * pTempPrefs;

WNDPROC OldKeyEditProc, OldCycleEditProc;
LRESULT CALLBACK NewKeyEditProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK NewCycleEditProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

void SubClassControl(HWND hWnd, CONTROLPROC NewProc)
{
	//OldKeyEditProc = (WNDPROC)GetWindowLong(hWnd, GWL_WNDPROC);
	//SetWindowLong(hWnd, GWL_WNDPROC, (LONG_PTR)NewProc);
    OldKeyEditProc = (WNDPROC)GetWindowLongPtr(hWnd, GWLP_WNDPROC);
    SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)NewProc);
}

LRESULT CALLBACK NewKeyEditProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WCHAR KeyName[128];

	switch(msg)
	{
	case WM_CHAR:
		return 0;
		break;
	case WM_GETDLGCODE:
		if (lParam)
		{
			if (((MSG*)lParam)->message == WM_KEYDOWN)
				return DLGC_WANTMESSAGE;
		}
		break;
	case WM_SYSKEYDOWN:
	case WM_KEYUP:
		if (wParam != VK_SNAPSHOT)
			return 0;
	case WM_KEYDOWN:
		switch (wParam)
		{
		case VK_SHIFT:
		case VK_CONTROL:
		case VK_MENU:
		case VK_LWIN:
		case VK_RWIN:
			return 0;
		}
		if (GetDlgCtrlID(hwnd) == IDC_EDIT_KEY)
			pTempPrefs->SetHotkeyCode(ComboBox_GetCurSel(GetDlgItem(GetParent(hwnd), IDC_COMBO1)), (UINT)wParam);
		else
			pTempPrefs->SetCycleKeyCode((UINT)wParam);
		GetKeyNameText((LONG)lParam, KeyName, 127);
		SetWindowText(hwnd, KeyName);
		SendMessage(hwnd, EM_SETSEL, MAKELONG(0x8000,0x0000), MAKELONG(0xffff,0xffff));
		return 0;
		break;
	}
	return CallWindowProc(OldKeyEditProc, hwnd, msg, wParam, lParam);
}

void InitComboBox(HWND hDlg)
{
	wstring deviceName;
	UINT count = pTempPrefs->GetCount();
	HWND hCombo = GetDlgItem(hDlg, IDC_COMBO1);

	ComboBox_ResetContent(hCombo);

	for (UINT i = 0; i < count; i++)
	{
		if (pTempPrefs->GetIsPresent(i) == false)
			break;
		ComboBox_AddString(hCombo, pTempPrefs->GetName(i, deviceName).c_str());
	}
	ComboBox_SetCurSel(hCombo, 0);
}

void ToggleHotkey(HWND hDlg)
{
	int device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
	if (device < 0)
		return;

	bool enabled = IsDlgButtonChecked(hDlg, IDC_CHECK_ENABLE_HOTKEY) == BST_CHECKED;
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_SHIFT_H), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CONTROL_H), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_ALT_H), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_WIN_H), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_EDIT_KEY), enabled);
	pTempPrefs->EnableHotkey(device, enabled);
	//if (enabled)
		//SetFocus(GetDlgItem(hDlg, IDC_EDIT_KEY));
}

void ToggleCycleKey(HWND hDlg)
{
	bool enabled = IsDlgButtonChecked(hDlg, IDC_CHECK_ENABLE_CYCLE) == BST_CHECKED;
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_SHIFT_C), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CONTROL_C), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_ALT_C), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_WIN_C), enabled);
	EnableWindow(GetDlgItem(hDlg, IDC_EDIT_KEY2), enabled);
	pTempPrefs->EnableCycleKey(enabled);
	//if (enabled)
		//SetFocus(GetDlgItem(hDlg, IDC_EDIT_KEY2));
}

void SetDialogItems(HWND hDlg)
{
	int device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));

	if (device < 0)
		return;

	wstring wstrS;
	UINT key, mods, check;

	pTempPrefs->GetHotkeyString(device, wstrS);
	key = pTempPrefs->GetHotkeyCode(device);
	mods = pTempPrefs->GetHotkeyMods(device);

	check = pTempPrefs->GetExcludeFromCycle(device) ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_CYCLE, check);

	check = pTempPrefs->GetIsHidden(device) ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_HIDE, check);

	check = (mods & MOD_SHIFT) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_SHIFT_H, check);

	check = (mods & MOD_CONTROL) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_CONTROL_H, check);

	check = (mods & MOD_ALT) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_ALT_H, check);
	
	check = (mods & MOD_WIN) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_WIN_H, check);
	

	if (key != 0)
			SetDlgItemText(hDlg, IDC_EDIT_KEY, wstrS.c_str());
	else
			SetDlgItemText(hDlg, IDC_EDIT_KEY, _T(""));

	key = pTempPrefs->GetCycleKeyCode();
	mods = pTempPrefs->GetCycleKeyMods();
	
	check = (mods & MOD_SHIFT) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_SHIFT_C, check);

	check = (mods & MOD_CONTROL) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_CONTROL_C, check);

	check = (mods & MOD_ALT) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_ALT_C, check);
	
	check = (mods & MOD_WIN) != 0 ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_WIN_C, check);
	
	if (key != 0)
	{
		pTempPrefs->GetCycleKeyString(wstrS);
		SetDlgItemText(hDlg, IDC_EDIT_KEY2, wstrS.c_str());
	}
	else
		SetDlgItemText(hDlg, IDC_EDIT_KEY2, _T(""));

	check = pTempPrefs->GetExcludeFromCycle(device) ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_CYCLE, check);
	EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CYCLE), !pTempPrefs->GetIsHidden(device));

	check = pTempPrefs->GetHotkeyEnabled(device) ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_ENABLE_HOTKEY, check);
	check = pTempPrefs->GetCycleKeyEnabled() ? BST_CHECKED : BST_UNCHECKED;
	CheckDlgButton(hDlg, IDC_CHECK_ENABLE_CYCLE, check);

	ToggleHotkey(hDlg);
	ToggleCycleKey(hDlg);
}

void HandleCycleKeyModClick(HWND hDlg, UINT param)
{
	UINT cyclemods = pTempPrefs->GetCycleKeyMods();

	switch (param)
	{
	case IDC_CHECK_SHIFT_C:
		cyclemods = IsDlgButtonChecked(hDlg, IDC_CHECK_SHIFT_C) == BST_CHECKED ? cyclemods | MOD_SHIFT : cyclemods & ~MOD_SHIFT;
		break;
	case IDC_CHECK_CONTROL_C:
		cyclemods = IsDlgButtonChecked(hDlg, IDC_CHECK_CONTROL_C) == BST_CHECKED ? cyclemods | MOD_CONTROL: cyclemods & ~MOD_CONTROL;
		break;
	case IDC_CHECK_ALT_C:
		cyclemods = IsDlgButtonChecked(hDlg, IDC_CHECK_ALT_C) == BST_CHECKED ? cyclemods | MOD_ALT : cyclemods & ~MOD_ALT;
		break;
	case IDC_CHECK_WIN_C:
		cyclemods = IsDlgButtonChecked(hDlg, IDC_CHECK_WIN_C) == BST_CHECKED ? cyclemods | MOD_WIN : cyclemods & ~MOD_WIN;
		break;
	}
	pTempPrefs->SetCycleKeyMods(cyclemods);
	//SetFocus(GetDlgItem(hDlg, IDC_EDIT_KEY2));

}

void HandleHotkeyModClick(HWND hDlg, UINT param)
{
	int device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
	if (device < 0)
		return;

	UINT hkmods = pTempPrefs->GetHotkeyMods(device);

	switch (param)
	{
	case IDC_CHECK_SHIFT_H:
		hkmods = IsDlgButtonChecked(hDlg, IDC_CHECK_SHIFT_H) == BST_CHECKED ? hkmods | MOD_SHIFT : hkmods & ~MOD_SHIFT;
		break;
	case IDC_CHECK_CONTROL_H:
		hkmods = IsDlgButtonChecked(hDlg, IDC_CHECK_CONTROL_H) == BST_CHECKED ? hkmods | MOD_CONTROL: hkmods & ~MOD_CONTROL;
		break;
	case IDC_CHECK_ALT_H:
		hkmods = IsDlgButtonChecked(hDlg, IDC_CHECK_ALT_H) == BST_CHECKED ? hkmods | MOD_ALT : hkmods & ~MOD_ALT;
		break;
	case IDC_CHECK_WIN_H:
		hkmods = IsDlgButtonChecked(hDlg, IDC_CHECK_WIN_H) == BST_CHECKED ? hkmods | MOD_WIN : hkmods & ~MOD_WIN;
		break;
	case IDC_CHECK_CYCLE:
		pTempPrefs->SetExcludeFromCycle(device, IsDlgButtonChecked(hDlg, IDC_CHECK_CYCLE) == BST_CHECKED); 
		break;
	case IDC_CHECK_HIDE:
		pTempPrefs->SetIsHidden(device, IsDlgButtonChecked(hDlg, IDC_CHECK_HIDE) == BST_CHECKED); 
		break;
	}
	pTempPrefs->SetHotkeyMods(device, hkmods);
	//SetFocus(GetDlgItem(hDlg, IDC_EDIT_KEY));
}

void HandleEditHotkeyChange(HWND hDlg)
{
	WCHAR keyname[256];
	int device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
	if (device < 0)
		return;

	GetDlgItemText(hDlg, IDC_EDIT_KEY, keyname, 255);
	pTempPrefs->SetHotkeyString(device, keyname);
}

void HandleEditCycleKeyChange(HWND hDlg)
{
	WCHAR keyname[256];

	GetDlgItemText(hDlg, IDC_EDIT_KEY2, keyname, 255);
	pTempPrefs->SetCycleKeyString(keyname);
}

INT_PTR CALLBACK SettingsDialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	static bool bBusy = false;

	if (bBusy)
		return (INT_PTR)TRUE;

	switch(uMsg)
	{
	case WM_INITDIALOG:
		bBusy = true;
		gOpenWindow = hDlg;
		pTempPrefs = reinterpret_cast<CQSESPrefs *> (lParam);
		SubClassControl(GetDlgItem(hDlg, IDC_EDIT_KEY), NewKeyEditProc);
		SubClassControl(GetDlgItem(hDlg, IDC_EDIT_KEY2), NewKeyEditProc);
		InitComboBox(hDlg);
		SetDialogItems(hDlg);
		CheckDlgButton(hDlg, IDC_CHECK_AUTOSTART, CheckAutoStart());
		bBusy = false;
		return (INT_PTR)TRUE;
	case WM_COMMAND:
		switch(LOWORD(wParam))
		{
		case IDC_CHECK_AUTOSTART:
			if (IsDlgButtonChecked(hDlg, IDC_CHECK_AUTOSTART) == TRUE)
				EnableAutoStart();
			else
				DisableAutoStart();
			return (INT_PTR)TRUE;
		case IDC_COMBO1:
			if(HIWORD(wParam) == CBN_SELCHANGE)
			{
				bBusy = true;
				SetDialogItems(hDlg);
				bBusy = false;
			}
			return (INT_PTR)TRUE;
		case IDCANCEL:
			EndDialog(hDlg, IDCANCEL);
			return (INT_PTR)TRUE;
		case IDC_BUTTON_DELETE:
			bBusy = true;
			pTempPrefs->Clear(ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1)));
			SetDialogItems(hDlg);
			bBusy = false;
			return (INT_PTR)TRUE;
		case IDOK:
			EndDialog(hDlg, IDOK);
			return (INT_PTR)TRUE;
		case IDC_CHECK_ENABLE_HOTKEY:
			bBusy = true;
			ToggleHotkey(hDlg);
			bBusy = false;
			return (INT_PTR)TRUE;
		case IDC_CHECK_ENABLE_CYCLE:
			bBusy = true;
			ToggleCycleKey(hDlg);
			bBusy = false;
			return (INT_PTR)TRUE;
		case IDC_CHECK_CYCLE:
			{
				bool enabled = IsDlgButtonChecked(hDlg, IDC_CHECK_CYCLE) == BST_CHECKED;
				pTempPrefs->SetExcludeFromCycle(ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1)), enabled);
				return (INT_PTR)TRUE;
			}
		case IDC_CHECK_HIDE:
			{
				bBusy = true;
				UINT checked = IsDlgButtonChecked(hDlg, IDC_CHECK_HIDE);
				UINT device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
				EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CYCLE), checked != BST_CHECKED);
				pTempPrefs->SetIsHidden(device, checked == BST_CHECKED);
				bBusy = false;
				return (INT_PTR)TRUE;
			}
		case IDC_EDIT_KEY:
			if (HIWORD(wParam) == EN_CHANGE)
			{
				bBusy = true;
				HandleEditHotkeyChange(hDlg);
				bBusy = false;
			}
			return (INT_PTR)TRUE;
		case IDC_EDIT_KEY2:
			if (HIWORD(wParam) == EN_CHANGE)
			{
				bBusy = true;
				HandleEditCycleKeyChange(hDlg);
				bBusy = false;
			}
			return (INT_PTR)TRUE;
		case IDC_CHECK_SHIFT_H:
		case IDC_CHECK_CONTROL_H:
		case IDC_CHECK_ALT_H:
		case IDC_CHECK_WIN_H:
				bBusy = true;
				HandleHotkeyModClick(hDlg, LOWORD(wParam));
				bBusy = false;
				return (INT_PTR)TRUE;
		case IDC_CHECK_SHIFT_C:
		case IDC_CHECK_CONTROL_C:
		case IDC_CHECK_ALT_C:
		case IDC_CHECK_WIN_C:
				bBusy = true;
				HandleCycleKeyModClick(hDlg, LOWORD(wParam));
				bBusy = false;
				return (INT_PTR)TRUE;
		default:
			break;
		}
		break;
	case WM_DESTROY:
		gOpenWindow = 0;
		return (INT_PTR)TRUE;
	}

	return (INT_PTR)FALSE;
}
