#pragma once

#include <string>
#include "windows.h"
#include <map>

const int cMaxDevices = 20;

enum prefs_sections {
	globals,
	device,
	none
};
using namespace std;

struct DevicePrefs
{
	UINT KeyMods;
	UINT KeyCode;
	std::wstring Name;
	std::wstring DeviceID;
	std::wstring HotkeyString;
	bool HasHotkey;
	bool IsExcludedFromCycle;
	bool IsHidden;
	bool IsPresent;

	DevicePrefs() :
		KeyMods(0),
		KeyCode(0),
		Name(),
		DeviceID(),
		HotkeyString(),
		HasHotkey(false),
		IsExcludedFromCycle(false),
		IsHidden(false),
		IsPresent(false)
	{}
};

class CQSESPrefs
{

public:

	// class config
	void Add(DevicePrefs *sdi);
	void Update(DevicePrefs *sdi);
	bool SetMax(int max);
	int GetMax() const { return mMax; };
	int GetCount() const { return mNext; };
	// hotkey
	void SetHotkeyString(int index, const wstring & keystring) { if (index < mNext) mDevices[index].HotkeyString = keystring; }
	wstring& GetHotkeyString(int index, wstring & keyString);
	void SetHotkeyCode(int index, UINT code);
	UINT GetHotkeyCode(int index) { if (index < mNext) return mDevices[index].KeyCode; else return 0; }
	void SetHotkeyMods(int index, UINT mods);
	UINT GetHotkeyMods(int index);
	void EnableHotkey(int index, bool enabled) { if (index < mNext) mDevices[index].HasHotkey = enabled; }
	bool GetHotkeyEnabled(int index);
	// cycle
	void SetExcludeFromCycle(int index, bool b) { if (index < mNext) mDevices[index].IsExcludedFromCycle = b; }
	bool GetExcludeFromCycle(int index);
	void SetCycleKeyString(const WCHAR * s) { mCycleKeyString = s; }
	wstring& GetCycleKeyString(wstring& s) const { return s = mCycleKeyString; }
	void SetCycleKeyCode(UINT key) { mKeyCode = key; }
	UINT GetCycleKeyCode() const { return mKeyCode; }
	void SetCycleKeyMods(UINT mods) { mKeyMods = mods; }
	UINT GetCycleKeyMods() const { return mKeyMods; }
	void EnableCycleKey(bool enabled) { mIsCycleKeyEnabled = enabled; }
	bool GetCycleKeyEnabled() const { return mIsCycleKeyEnabled; }
	// misc
	void SetIsHidden(int index, bool b) { if (index < mNext) mDevices[index].IsHidden = b; }
	bool GetIsHidden(int index);
	void SetIsPresent(int index, bool b) { if (index < mNext) mDevices[index].IsPresent = b; }
	bool GetIsPresent(int index);
	void Swap(int d1, int d2);
	void Sort();
	// device info
	bool GetByName(wstring& name, DevicePrefs * sdi);
	int FindByID(const wstring& id);
	wstring& GetName(int index, wstring & s);
	wstring& GetID(int index, wstring& s);
	// prefs class
	bool Init(int max);
	void Clear(int index);
	void ResetPresent();
	// load & save
	bool Load();
	bool Save();

	CQSESPrefs( const CQSESPrefs& other );
	CQSESPrefs() : mDevices(0), mNext(0), mMax(cMaxDevices), mKeyCode(0),
							mKeyMods(0), mCycleKeyString(L""), mIsCycleKeyEnabled(false)
							{ }
	~CQSESPrefs() { delete[] mDevices; }
	CQSESPrefs& operator=(const CQSESPrefs& source);

private:

	void Erase(int index);
	void Remove(int index);
	void RemoveDupes();
	int FindByName(wstring & name);
	bool ReadConfig(const wstring &fileName, map <wstring, map <wstring, wstring> > & sections);

	DevicePrefs * mDevices;
	int mMax;
	int mNext;
	UINT mKeyCode;
	UINT mKeyMods;
	wstring mCycleKeyString;
	bool mIsCycleKeyEnabled;
	static const int mLimit = cMaxDevices;

};
