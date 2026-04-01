#pragma once

#include <string>
#include <vector>
#include <map>
#include "windows.h"

inline constexpr int cMaxDevices = 20;

struct DevicePrefs
{
    UINT        KeyMods               = 0;
    UINT        KeyCode               = 0;
    std::wstring Name;           // System-provided friendly name (from Windows)
    std::wstring CustomName;     // User-specified display name (overrides Name when non-empty)
    std::wstring DeviceID;
    std::wstring HotkeyString;
    bool        HasHotkey             = false;
    bool        IsExcludedFromCycle   = false;
    bool        IsHidden              = false;
    bool        IsPresent             = false;
};

class CQSESPrefs
{
public:
    CQSESPrefs() = default;
    CQSESPrefs(const CQSESPrefs&) = default;
    CQSESPrefs(CQSESPrefs&&) = default;
    CQSESPrefs& operator=(const CQSESPrefs&) = default;
    CQSESPrefs& operator=(CQSESPrefs&&) = default;
    ~CQSESPrefs() = default;

    // Collection management
    void Add(const DevicePrefs& sdi);
    void Update(const DevicePrefs& sdi);
    int  GetCount() const { return static_cast<int>(mDevices.size()); }

    // Per-device hotkey
    void         SetHotkeyString(int index, const std::wstring& keystring);
    std::wstring GetHotkeyString(int index) const;
    void         SetHotkeyCode(int index, UINT code);
    UINT         GetHotkeyCode(int index) const;
    void         SetHotkeyMods(int index, UINT mods);
    UINT         GetHotkeyMods(int index) const;
    void         EnableHotkey(int index, bool enabled);
    bool         GetHotkeyEnabled(int index) const;

    // Cycle key
    void         SetExcludeFromCycle(int index, bool b);
    bool         GetExcludeFromCycle(int index) const;
    void         SetCycleKeyString(const std::wstring& s)    { mCycleKeyString = s; }
    std::wstring GetCycleKeyString() const                   { return mCycleKeyString; }
    void         SetCycleKeyCode(UINT key)                   { mKeyCode = key; }
    UINT         GetCycleKeyCode() const                     { return mKeyCode; }
    void         SetCycleKeyMods(UINT mods)                  { mKeyMods = mods; }
    UINT         GetCycleKeyMods() const                     { return mKeyMods; }
    void         EnableCycleKey(bool enabled)                { mIsCycleKeyEnabled = enabled; }
    bool         GetCycleKeyEnabled() const                  { return mIsCycleKeyEnabled; }

    // Per-device visibility / presence
    void SetIsHidden(int index, bool b);
    bool GetIsHidden(int index) const;
    void SetIsPresent(int index, bool b);
    bool GetIsPresent(int index) const;

    // Ordering
    void Swap(int d1, int d2);
    void Sort();

    // Per-device custom name
    void         SetCustomName(int index, const std::wstring& name);
    std::wstring GetCustomName(int index) const;

    // Lookup
    int          FindByID(const std::wstring& id) const;
    std::wstring GetName(int index) const;        // returns system name (from Windows)
    std::wstring GetDisplayName(int index) const; // returns CustomName if set, else Name
    std::wstring GetID(int index) const;

    // Lifecycle
    void Clear(int index);      // clears hotkey/flags but keeps name/ID
    void ResetPresent();

    // Persistence
    bool Load();
    bool Save();

private:
    void RemoveDupes();
    void Remove(int index);
    void Erase(int index);
    int  FindByName(const std::wstring& name) const;
    bool ReadConfig(const std::wstring& fileName,
                    std::map<std::wstring, std::map<std::wstring, std::wstring>>& sections) const;

    std::vector<DevicePrefs> mDevices;
    UINT         mKeyCode           = 0;
    UINT         mKeyMods           = 0;
    std::wstring mCycleKeyString;
    bool         mIsCycleKeyEnabled = false;
};
