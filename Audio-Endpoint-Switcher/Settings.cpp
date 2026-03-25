#include <Windowsx.h>
#include <string>
#include "resource.h"
#include "Settings.h"

using namespace std;

extern HWND  gOpenWindow;
extern void  EnableAutoStart();
extern void  DisableAutoStart();
extern UINT  CheckAutoStart();

static CQSESPrefs* pTempPrefs = nullptr;

// ---------------------------------------------------------------------------
// Each subclassed edit control needs its own saved WNDPROC.
// ---------------------------------------------------------------------------
static WNDPROC OldKeyEditProc   = nullptr;
static WNDPROC OldCycleEditProc = nullptr;

// ---------------------------------------------------------------------------
// Shared key-capture logic; called by both subclass procs.
// ---------------------------------------------------------------------------
static LRESULT HandleKeyCapture(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                 WNDPROC oldProc, bool isCycleKey)
{
    switch (msg)
    {
    case WM_CHAR:
        return 0;

    case WM_GETDLGCODE:
        if (lParam)
        {
            const MSG* pMsg = reinterpret_cast<const MSG*>(lParam);
            if (pMsg->message == WM_KEYDOWN)
                return DLGC_WANTMESSAGE;
        }
        break;

    case WM_SYSKEYDOWN:
    case WM_KEYUP:
        if (wParam != VK_SNAPSHOT)
            return 0;
        [[fallthrough]];

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
        {
            WCHAR keyName[128] = {};
            if (isCycleKey)
            {
                pTempPrefs->SetCycleKeyCode(static_cast<UINT>(wParam));
            }
            else
            {
                int device = ComboBox_GetCurSel(GetDlgItem(GetParent(hwnd), IDC_COMBO1));
                pTempPrefs->SetHotkeyCode(device, static_cast<UINT>(wParam));
            }
            GetKeyNameTextW(static_cast<LONG>(lParam), keyName, 127);
            SetWindowTextW(hwnd, keyName);
            SendMessageW(hwnd, EM_SETSEL, 0, -1);
        }
        return 0;
    }
    return CallWindowProcW(oldProc, hwnd, msg, wParam, lParam);
}

static LRESULT CALLBACK NewKeyEditProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    return HandleKeyCapture(hwnd, msg, wParam, lParam, OldKeyEditProc, false);
}

static LRESULT CALLBACK NewCycleEditProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    return HandleKeyCapture(hwnd, msg, wParam, lParam, OldCycleEditProc, true);
}

// ---------------------------------------------------------------------------
// Dialog helpers
// ---------------------------------------------------------------------------

static void InitComboBox(HWND hDlg)
{
    const int count   = pTempPrefs->GetCount();
    HWND      hCombo  = GetDlgItem(hDlg, IDC_COMBO1);

    ComboBox_ResetContent(hCombo);
    for (int i = 0; i < count; ++i)
    {
        if (!pTempPrefs->GetIsPresent(i))
            break;
        ComboBox_AddString(hCombo, pTempPrefs->GetName(i).c_str());
    }
    ComboBox_SetCurSel(hCombo, 0);
}

static void ToggleHotkey(HWND hDlg)
{
    int device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
    if (device < 0) return;

    bool enabled = IsDlgButtonChecked(hDlg, IDC_CHECK_ENABLE_HOTKEY) == BST_CHECKED;
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_SHIFT_H),   enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CONTROL_H), enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_ALT_H),     enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_WIN_H),     enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_EDIT_KEY),        enabled);
    pTempPrefs->EnableHotkey(device, enabled);
}

static void ToggleCycleKey(HWND hDlg)
{
    bool enabled = IsDlgButtonChecked(hDlg, IDC_CHECK_ENABLE_CYCLE) == BST_CHECKED;
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_SHIFT_C),   enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CONTROL_C), enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_ALT_C),     enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_WIN_C),     enabled);
    EnableWindow(GetDlgItem(hDlg, IDC_EDIT_KEY2),       enabled);
    pTempPrefs->EnableCycleKey(enabled);
}

static void SetDialogItems(HWND hDlg)
{
    int device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
    if (device < 0) return;

    // Per-device hotkey controls
    UINT mods = pTempPrefs->GetHotkeyMods(device);
    UINT key  = pTempPrefs->GetHotkeyCode(device);

    CheckDlgButton(hDlg, IDC_CHECK_CYCLE,  pTempPrefs->GetExcludeFromCycle(device) ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_HIDE,   pTempPrefs->GetIsHidden(device)         ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_SHIFT_H,   (mods & MOD_SHIFT)   ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_CONTROL_H, (mods & MOD_CONTROL) ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_ALT_H,     (mods & MOD_ALT)     ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_WIN_H,     (mods & MOD_WIN)     ? BST_CHECKED : BST_UNCHECKED);

    if (key != 0)
        SetDlgItemTextW(hDlg, IDC_EDIT_KEY, pTempPrefs->GetHotkeyString(device).c_str());
    else
        SetDlgItemTextW(hDlg, IDC_EDIT_KEY, L"");

    // Cycle key controls
    UINT cmods = pTempPrefs->GetCycleKeyMods();
    UINT ckey  = pTempPrefs->GetCycleKeyCode();

    CheckDlgButton(hDlg, IDC_CHECK_SHIFT_C,   (cmods & MOD_SHIFT)   ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_CONTROL_C, (cmods & MOD_CONTROL) ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_ALT_C,     (cmods & MOD_ALT)     ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_WIN_C,     (cmods & MOD_WIN)     ? BST_CHECKED : BST_UNCHECKED);

    if (ckey != 0)
        SetDlgItemTextW(hDlg, IDC_EDIT_KEY2, pTempPrefs->GetCycleKeyString().c_str());
    else
        SetDlgItemTextW(hDlg, IDC_EDIT_KEY2, L"");

    EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CYCLE), !pTempPrefs->GetIsHidden(device));

    CheckDlgButton(hDlg, IDC_CHECK_ENABLE_HOTKEY,
                   pTempPrefs->GetHotkeyEnabled(device)   ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hDlg, IDC_CHECK_ENABLE_CYCLE,
                   pTempPrefs->GetCycleKeyEnabled()       ? BST_CHECKED : BST_UNCHECKED);

    ToggleHotkey(hDlg);
    ToggleCycleKey(hDlg);
}

static void HandleCycleKeyModClick(HWND hDlg, UINT param)
{
    UINT mods = pTempPrefs->GetCycleKeyMods();
    auto toggle = [&](UINT id, UINT flag)
    {
        mods = IsDlgButtonChecked(hDlg, id) == BST_CHECKED ? mods | flag : mods & ~flag;
    };
    switch (param)
    {
    case IDC_CHECK_SHIFT_C:   toggle(IDC_CHECK_SHIFT_C,   MOD_SHIFT);   break;
    case IDC_CHECK_CONTROL_C: toggle(IDC_CHECK_CONTROL_C, MOD_CONTROL); break;
    case IDC_CHECK_ALT_C:     toggle(IDC_CHECK_ALT_C,     MOD_ALT);     break;
    case IDC_CHECK_WIN_C:     toggle(IDC_CHECK_WIN_C,     MOD_WIN);     break;
    }
    pTempPrefs->SetCycleKeyMods(mods);
}

static void HandleHotkeyModClick(HWND hDlg, UINT param)
{
    int device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
    if (device < 0) return;

    UINT mods = pTempPrefs->GetHotkeyMods(device);
    auto toggle = [&](UINT id, UINT flag)
    {
        mods = IsDlgButtonChecked(hDlg, id) == BST_CHECKED ? mods | flag : mods & ~flag;
    };
    switch (param)
    {
    case IDC_CHECK_SHIFT_H:   toggle(IDC_CHECK_SHIFT_H,   MOD_SHIFT);   break;
    case IDC_CHECK_CONTROL_H: toggle(IDC_CHECK_CONTROL_H, MOD_CONTROL); break;
    case IDC_CHECK_ALT_H:     toggle(IDC_CHECK_ALT_H,     MOD_ALT);     break;
    case IDC_CHECK_WIN_H:     toggle(IDC_CHECK_WIN_H,     MOD_WIN);     break;
    case IDC_CHECK_CYCLE:
        pTempPrefs->SetExcludeFromCycle(device,
            IsDlgButtonChecked(hDlg, IDC_CHECK_CYCLE) == BST_CHECKED);
        return; // mods unchanged
    case IDC_CHECK_HIDE:
        pTempPrefs->SetIsHidden(device,
            IsDlgButtonChecked(hDlg, IDC_CHECK_HIDE) == BST_CHECKED);
        return; // mods unchanged
    }
    pTempPrefs->SetHotkeyMods(device, mods);
}

// ---------------------------------------------------------------------------
// Dialog procedure
// ---------------------------------------------------------------------------

INT_PTR CALLBACK SettingsDialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    static bool bBusy = false;
    if (bBusy) return TRUE;

    switch (uMsg)
    {
    case WM_INITDIALOG:
    {
        bBusy      = true;
        gOpenWindow = hDlg;
        pTempPrefs  = reinterpret_cast<CQSESPrefs*>(lParam);

        // Subclass the two key-capture edit boxes with separate procs.
        OldKeyEditProc   = reinterpret_cast<WNDPROC>(
            SetWindowLongPtrW(GetDlgItem(hDlg, IDC_EDIT_KEY),  GWLP_WNDPROC,
                              reinterpret_cast<LONG_PTR>(NewKeyEditProc)));
        OldCycleEditProc = reinterpret_cast<WNDPROC>(
            SetWindowLongPtrW(GetDlgItem(hDlg, IDC_EDIT_KEY2), GWLP_WNDPROC,
                              reinterpret_cast<LONG_PTR>(NewCycleEditProc)));

        InitComboBox(hDlg);
        SetDialogItems(hDlg);
        CheckDlgButton(hDlg, IDC_CHECK_AUTOSTART, CheckAutoStart());
        bBusy = false;
        return TRUE;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_CHECK_AUTOSTART:
            if (IsDlgButtonChecked(hDlg, IDC_CHECK_AUTOSTART) == BST_CHECKED)
                EnableAutoStart();
            else
                DisableAutoStart();
            return TRUE;

        case IDC_COMBO1:
            if (HIWORD(wParam) == CBN_SELCHANGE)
            {
                bBusy = true;
                SetDialogItems(hDlg);
                bBusy = false;
            }
            return TRUE;

        case IDCANCEL:
            EndDialog(hDlg, IDCANCEL);
            return TRUE;

        case IDC_BUTTON_DELETE:
            bBusy = true;
            pTempPrefs->Clear(ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1)));
            SetDialogItems(hDlg);
            bBusy = false;
            return TRUE;

        case IDOK:
            EndDialog(hDlg, IDOK);
            return TRUE;

        case IDC_CHECK_ENABLE_HOTKEY:
            bBusy = true;
            ToggleHotkey(hDlg);
            bBusy = false;
            return TRUE;

        case IDC_CHECK_ENABLE_CYCLE:
            bBusy = true;
            ToggleCycleKey(hDlg);
            bBusy = false;
            return TRUE;

        case IDC_CHECK_CYCLE:
            pTempPrefs->SetExcludeFromCycle(
                ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1)),
                IsDlgButtonChecked(hDlg, IDC_CHECK_CYCLE) == BST_CHECKED);
            return TRUE;

        case IDC_CHECK_HIDE:
        {
            bBusy = true;
            UINT device  = static_cast<UINT>(ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1)));
            bool checked = IsDlgButtonChecked(hDlg, IDC_CHECK_HIDE) == BST_CHECKED;
            EnableWindow(GetDlgItem(hDlg, IDC_CHECK_CYCLE), !checked);
            pTempPrefs->SetIsHidden(device, checked);
            bBusy = false;
            return TRUE;
        }

        case IDC_EDIT_KEY:
            if (HIWORD(wParam) == EN_CHANGE)
            {
                bBusy = true;
                WCHAR keyname[256] = {};
                int   device = ComboBox_GetCurSel(GetDlgItem(hDlg, IDC_COMBO1));
                GetDlgItemTextW(hDlg, IDC_EDIT_KEY, keyname, 255);
                pTempPrefs->SetHotkeyString(device, keyname);
                bBusy = false;
            }
            return TRUE;

        case IDC_EDIT_KEY2:
            if (HIWORD(wParam) == EN_CHANGE)
            {
                bBusy = true;
                WCHAR keyname[256] = {};
                GetDlgItemTextW(hDlg, IDC_EDIT_KEY2, keyname, 255);
                pTempPrefs->SetCycleKeyString(keyname);
                bBusy = false;
            }
            return TRUE;

        case IDC_CHECK_SHIFT_H:
        case IDC_CHECK_CONTROL_H:
        case IDC_CHECK_ALT_H:
        case IDC_CHECK_WIN_H:
            bBusy = true;
            HandleHotkeyModClick(hDlg, LOWORD(wParam));
            bBusy = false;
            return TRUE;

        case IDC_CHECK_SHIFT_C:
        case IDC_CHECK_CONTROL_C:
        case IDC_CHECK_ALT_C:
        case IDC_CHECK_WIN_C:
            bBusy = true;
            HandleCycleKeyModClick(hDlg, LOWORD(wParam));
            bBusy = false;
            return TRUE;

        default:
            break;
        }
        break;

    case WM_DESTROY:
        gOpenWindow = nullptr;
        return TRUE;
    }
    return FALSE;
}
