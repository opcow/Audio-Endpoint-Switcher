#include <Shlobj.h>
#include <windows.h>
#include <wchar.h>
#include <algorithm>
#include <regex>
#include <vector>
#include <cstdio>

#include "Prefs.h"

using namespace std;

// ---------------------------------------------------------------------------
// String helpers
// ---------------------------------------------------------------------------

static void ltrim(wstring& s)
{
    s.erase(s.begin(),
            find_if(s.begin(), s.end(), [](wchar_t ch) { return !iswspace(ch); }));
}

static void rtrim(wstring& s)
{
    s.erase(find_if(s.rbegin(), s.rend(), [](wchar_t ch) { return !iswspace(ch); }).base(),
            s.end());
}

static void trim(wstring& s) { ltrim(s); rtrim(s); }

// ---------------------------------------------------------------------------
// Bounds-check helper
// ---------------------------------------------------------------------------
static bool ValidIndex(int index, int count) { return index >= 0 && index < count; }

// ---------------------------------------------------------------------------
// Collection management
// ---------------------------------------------------------------------------

// Add: replaces the *entire* DevicePrefs record on collision (hotkeys, flags,
// everything). Use when importing a fully-populated record.
void CQSESPrefs::Add(const DevicePrefs& sdi)
{
    int index = FindByID(sdi.DeviceID);
    if (index != -1)
    {
        mDevices[index] = sdi;
        mDevices[index].IsPresent = true;
    }
    else
    {
        if (static_cast<int>(mDevices.size()) >= cMaxDevices)
            Remove(0);
        DevicePrefs d = sdi;
        d.IsPresent = true;
        mDevices.push_back(std::move(d));
    }
}

// Update: refreshes only DeviceID, Name, and IsPresent on collision, leaving
// user-configured fields (hotkeys, hidden, excluded) untouched. Use when
// re-enumerating live audio endpoints.
void CQSESPrefs::Update(const DevicePrefs& sdi)
{
    int index = FindByID(sdi.DeviceID);
    if (index != -1)
    {
        mDevices[index].DeviceID  = sdi.DeviceID;
        mDevices[index].Name      = sdi.Name;
        mDevices[index].IsPresent = true;
    }
    else
    {
        if (static_cast<int>(mDevices.size()) >= cMaxDevices)
            Remove(0);
        DevicePrefs d = sdi;
        d.IsPresent = true;
        mDevices.push_back(std::move(d));
    }
}

// ---------------------------------------------------------------------------
// Per-device hotkey
// ---------------------------------------------------------------------------

void CQSESPrefs::SetHotkeyString(int index, const wstring& keystring)
{
    if (ValidIndex(index, GetCount()))
        mDevices[index].HotkeyString = keystring;
}

wstring CQSESPrefs::GetHotkeyString(int index) const
{
    if (ValidIndex(index, GetCount()))
        return mDevices[index].HotkeyString;
    return {};
}

void CQSESPrefs::SetHotkeyCode(int index, UINT code)
{
    if (!ValidIndex(index, GetCount())) return;
    mDevices[index].KeyCode = code;
    if (mDevices[index].KeyCode == 0 && mDevices[index].KeyMods == 0)
        mDevices[index].HasHotkey = false;
}

UINT CQSESPrefs::GetHotkeyCode(int index) const
{
    return ValidIndex(index, GetCount()) ? mDevices[index].KeyCode : 0u;
}

void CQSESPrefs::SetHotkeyMods(int index, UINT mods)
{
    if (!ValidIndex(index, GetCount())) return;
    mDevices[index].KeyMods = mods;
    if (mDevices[index].KeyCode == 0 && mDevices[index].KeyMods == 0)
        mDevices[index].HasHotkey = false;
}

UINT CQSESPrefs::GetHotkeyMods(int index) const
{
    return ValidIndex(index, GetCount()) ? mDevices[index].KeyMods : 0u;
}

void CQSESPrefs::EnableHotkey(int index, bool enabled)
{
    if (ValidIndex(index, GetCount()))
        mDevices[index].HasHotkey = enabled;
}

bool CQSESPrefs::GetHotkeyEnabled(int index) const
{
    return ValidIndex(index, GetCount()) && mDevices[index].HasHotkey;
}

// ---------------------------------------------------------------------------
// Cycle / visibility
// ---------------------------------------------------------------------------

void CQSESPrefs::SetExcludeFromCycle(int index, bool b)
{
    if (ValidIndex(index, GetCount()))
        mDevices[index].IsExcludedFromCycle = b;
}

bool CQSESPrefs::GetExcludeFromCycle(int index) const
{
    return ValidIndex(index, GetCount()) && mDevices[index].IsExcludedFromCycle;
}

void CQSESPrefs::SetIsHidden(int index, bool b)
{
    if (ValidIndex(index, GetCount()))
        mDevices[index].IsHidden = b;
}

bool CQSESPrefs::GetIsHidden(int index) const
{
    return ValidIndex(index, GetCount()) && mDevices[index].IsHidden;
}

void CQSESPrefs::SetIsPresent(int index, bool b)
{
    if (ValidIndex(index, GetCount()))
        mDevices[index].IsPresent = b;
}

bool CQSESPrefs::GetIsPresent(int index) const
{
    return ValidIndex(index, GetCount()) && mDevices[index].IsPresent;
}

// ---------------------------------------------------------------------------
// Ordering
// ---------------------------------------------------------------------------

void CQSESPrefs::Swap(int d1, int d2)
{
    if (ValidIndex(d1, GetCount()) && ValidIndex(d2, GetCount()))
        std::swap(mDevices[d1], mDevices[d2]);
}

void CQSESPrefs::Sort()
{
    // Move all present devices to the front, preserving relative order.
    std::stable_partition(mDevices.begin(), mDevices.end(),
                          [](const DevicePrefs& d) { return d.IsPresent; });
}

// ---------------------------------------------------------------------------
// Lookup
// ---------------------------------------------------------------------------

int CQSESPrefs::FindByID(const wstring& id) const
{
    for (int i = 0; i < GetCount(); ++i)
        if (mDevices[i].DeviceID == id)
            return i;
    return -1;
}

int CQSESPrefs::FindByName(const wstring& name) const
{
    for (int i = 0; i < GetCount(); ++i)
        if (mDevices[i].Name == name)
            return i;
    return -1;
}

wstring CQSESPrefs::GetName(int index) const
{
    return ValidIndex(index, GetCount()) ? mDevices[index].Name : wstring{};
}

void CQSESPrefs::SetCustomName(int index, const wstring& name)
{
    if (ValidIndex(index, GetCount()))
        mDevices[index].CustomName = name;
}

wstring CQSESPrefs::GetCustomName(int index) const
{
    return ValidIndex(index, GetCount()) ? mDevices[index].CustomName : wstring{};
}

wstring CQSESPrefs::GetDisplayName(int index) const
{
    if (!ValidIndex(index, GetCount())) return {};
    const auto& custom = mDevices[index].CustomName;
    return custom.empty() ? mDevices[index].Name : custom;
}

wstring CQSESPrefs::GetID(int index) const
{
    return ValidIndex(index, GetCount()) ? mDevices[index].DeviceID : wstring{};
}

// ---------------------------------------------------------------------------
// Lifecycle
// ---------------------------------------------------------------------------

void CQSESPrefs::Clear(int index)
{
    if (!ValidIndex(index, GetCount())) return;
    mDevices[index].HotkeyString.clear();
    mDevices[index].CustomName.clear();
    mDevices[index].KeyCode             = 0;
    mDevices[index].KeyMods             = 0;
    mDevices[index].HasHotkey           = false;
    mDevices[index].IsExcludedFromCycle = false;
    mDevices[index].IsPresent           = false;
    mDevices[index].IsHidden            = false;
}

void CQSESPrefs::Erase(int index)
{
    if (!ValidIndex(index, GetCount())) return;
    mDevices[index] = DevicePrefs{};
}

void CQSESPrefs::ResetPresent()
{
    for (auto& d : mDevices)
        d.IsPresent = false;
}

void CQSESPrefs::Remove(int index)
{
    if (!ValidIndex(index, GetCount())) return;
    mDevices.erase(mDevices.begin() + index);
}

void CQSESPrefs::RemoveDupes()
{
    for (int i = 0; i < GetCount() - 1; ++i)
        for (int j = i + 1; j < GetCount(); ++j)
            if (mDevices[i].DeviceID == mDevices[j].DeviceID)
            {
                Remove(j);
                --j;
            }
}

// ---------------------------------------------------------------------------
// Persistence
// ---------------------------------------------------------------------------

bool CQSESPrefs::Save()
{
    RemoveDupes();

    LPWSTR wstrPrefsPath = nullptr;
    SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, nullptr, &wstrPrefsPath);
    wstring filePath(wstrPrefsPath);
    CoTaskMemFree(wstrPrefsPath);

    filePath += L"\\Audio-Endpoint-Switcher";
    CreateDirectoryW(filePath.c_str(), nullptr);
    filePath += L"\\Preferences.ini";

    FILE* fh = nullptr;
    if (_wfopen_s(&fh, filePath.c_str(), L"w") != 0 || fh == nullptr)
        return false;

    fwprintf(fh, L"[Globals]\n");
    fwprintf(fh, L"\tEnableCycling = %d\n",   mIsCycleKeyEnabled ? 1 : 0);
    fwprintf(fh, L"\tKeyCode = 0x%x\n",        mKeyCode);
    fwprintf(fh, L"\tKeyString = %s\n",         mCycleKeyString.c_str());
    fwprintf(fh, L"\tAltKey = %d\n",            (mKeyMods & MOD_ALT)     ? 1 : 0);
    fwprintf(fh, L"\tControlKey = %d\n",        (mKeyMods & MOD_CONTROL) ? 1 : 0);
    fwprintf(fh, L"\tShiftKey = %d\n",          (mKeyMods & MOD_SHIFT)   ? 1 : 0);
    fwprintf(fh, L"\tWinKey = %d\n",            (mKeyMods & MOD_WIN)     ? 1 : 0);

    for (const auto& dev : mDevices)
    {
        fwprintf(fh, L"[%s]\n",               dev.DeviceID.c_str());
        fwprintf(fh, L"\tName = %s\n",         dev.Name.c_str());
        fwprintf(fh, L"\tCustomName = %s\n",   dev.CustomName.c_str());
        fwprintf(fh, L"\tHasKey = %d\n",       dev.HasHotkey           ? 1 : 0);
        fwprintf(fh, L"\tKeyString = %s\n",    dev.HotkeyString.c_str());
        fwprintf(fh, L"\tExcluded = %d\n",     dev.IsExcludedFromCycle ? 1 : 0);
        fwprintf(fh, L"\tHidden = %d\n",       dev.IsHidden            ? 1 : 0);
        fwprintf(fh, L"\tPresent = %d\n",      dev.IsPresent           ? 1 : 0);
        fwprintf(fh, L"\tKeyCode = 0x%x\n",    dev.KeyCode);
        fwprintf(fh, L"\tAltKey = %d\n",       (dev.KeyMods & MOD_ALT)     ? 1 : 0);
        fwprintf(fh, L"\tControlKey = %d\n",   (dev.KeyMods & MOD_CONTROL) ? 1 : 0);
        fwprintf(fh, L"\tShiftKey = %d\n",     (dev.KeyMods & MOD_SHIFT)   ? 1 : 0);
        fwprintf(fh, L"\tWinKey = %d\n",       (dev.KeyMods & MOD_WIN)     ? 1 : 0);
    }
    fclose(fh);
    return true;
}

bool CQSESPrefs::ReadConfig(const wstring& fileName,
                             map<wstring, map<wstring, wstring>>& sections) const
{
    constexpr int cBuffLen = 1024;
    wchar_t wBuffer[cBuffLen];
    wstring sectionStr(L"[]");

    FILE* fh = nullptr;
    if (_wfopen_s(&fh, fileName.c_str(), L"r") != 0 || fh == nullptr)
        return false;

    while (fgetws(wBuffer, cBuffLen, fh))
    {
        wstring str(wBuffer);
        trim(str);
        if (str.empty() || str[0] == L'#')
            continue;

        // Strip inline comments
        size_t p = str.find(L'#');
        if (p != wstring::npos)
        {
            str.erase(p);
            rtrim(str);
        }

        if (str[0] == L'[')
        {
            transform(str.begin(), str.end(), str.begin(), ::towlower);
            sectionStr = str;
        }
        else
        {
            size_t eq = str.find(L'=');
            if (eq != wstring::npos)
            {
                wstring key = str.substr(0, eq);
                wstring val = str.substr(eq + 1);
                rtrim(key);
                ltrim(val);
                transform(key.begin(), key.end(), key.begin(), ::towlower);
                sections[sectionStr][key] = val;
            }
        }
    }
    fclose(fh);
    return true;
}

bool CQSESPrefs::Load()
{
    LPWSTR wstrPrefsPath = nullptr;
    SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, nullptr, &wstrPrefsPath);
    wstring filePath(wstrPrefsPath);
    CoTaskMemFree(wstrPrefsPath);

    filePath += L"\\Audio-Endpoint-Switcher\\Preferences.ini";

    map<wstring, map<wstring, wstring>> sections;
    if (!ReadConfig(filePath, sections))
        return false;

    // Helper lambdas
    auto getBool = [](const map<wstring, wstring>& sec, const wstring& key) -> bool
    {
        auto it = sec.find(key);
        return it != sec.end() && it->second == L"1";
    };
    auto getUInt = [](const map<wstring, wstring>& sec, const wstring& key) -> UINT
    {
        auto it = sec.find(key);
        return it != sec.end() ? static_cast<UINT>(wcstoul(it->second.c_str(), nullptr, 16)) : 0u;
    };
    auto getString = [](const map<wstring, wstring>& sec, const wstring& key) -> wstring
    {
        auto it = sec.find(key);
        return it != sec.end() ? it->second : wstring{};
    };

    // Globals section
    auto globalIt = sections.find(L"[globals]");
    if (globalIt != sections.end())
    {
        const auto& g = globalIt->second;
        mKeyCode           = getUInt(g, L"keycode");
        mIsCycleKeyEnabled = getBool(g, L"enablecycling");
        mCycleKeyString    = getString(g, L"keystring");
        mKeyMods           = 0;
        if (getBool(g, L"altkey"))     mKeyMods |= MOD_ALT;
        if (getBool(g, L"controlkey")) mKeyMods |= MOD_CONTROL;
        if (getBool(g, L"shiftkey"))   mKeyMods |= MOD_SHIFT;
        if (getBool(g, L"winkey"))     mKeyMods |= MOD_WIN;
    }

    // Device sections — keyed by Windows audio endpoint GUID pattern
    wregex guidRegex(L"\\[\\{\\d\\.\\d\\.\\d\\.\\d{8}\\}\\.\\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\\}\\]");
    wsmatch wMatch;

    mDevices.clear();
    for (const auto& kv : sections)
    {
        if (!regex_match(kv.first.cbegin(), kv.first.cend(), wMatch, guidRegex))
            continue;
        if (static_cast<int>(mDevices.size()) >= cMaxDevices)
            break;

        const auto& sec = kv.second;
        DevicePrefs dev;
        dev.DeviceID            = kv.first.substr(1, kv.first.size() - 2);
        dev.Name                = getString(sec, L"name");
        dev.CustomName          = getString(sec, L"customname");
        dev.HotkeyString        = getString(sec, L"keystring");
        dev.KeyCode             = getUInt(sec, L"keycode");
        dev.HasHotkey           = getBool(sec, L"haskey");
        dev.IsExcludedFromCycle = getBool(sec, L"excluded");
        dev.IsHidden            = getBool(sec, L"hidden");
        dev.IsPresent           = getBool(sec, L"present");
        dev.KeyMods             = 0;
        if (getBool(sec, L"altkey"))     dev.KeyMods |= MOD_ALT;
        if (getBool(sec, L"controlkey")) dev.KeyMods |= MOD_CONTROL;
        if (getBool(sec, L"shiftkey"))   dev.KeyMods |= MOD_SHIFT;
        if (getBool(sec, L"winkey"))     dev.KeyMods |= MOD_WIN;
        mDevices.push_back(std::move(dev));
    }
    return true;
}
