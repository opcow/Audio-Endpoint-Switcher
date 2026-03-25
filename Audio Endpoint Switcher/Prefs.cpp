#include <Shlobj.h>
#include <windows.h>
#include <wchar.h>
#include <algorithm>
#include <functional> 
#include <regex>

#include "Prefs.h"

using namespace std;

// trim from start (in place)
static inline void ltrim(std::wstring &s) {
	s.erase(s.begin(), std::find_if(s.begin(), s.end(),
		[](int ch) { return !iswspace(ch); }));
}

// trim from end (in place)
static inline void rtrim(std::wstring &s) {
	s.erase(std::find_if(s.rbegin(), s.rend(),
		[](int ch) { return !iswspace(ch); }).base(), s.end());
}

// trim from both ends (in place)
static inline void trim(std::wstring &s) {
	ltrim(s);
	rtrim(s);
}

CQSESPrefs::CQSESPrefs(const CQSESPrefs& other) :
	mMax(other.mMax),
	mNext(other.mNext),
	mKeyCode(other.mKeyCode),
	mKeyMods(other.mKeyMods),
	mCycleKeyString(other.mCycleKeyString),
	mIsCycleKeyEnabled(other.mIsCycleKeyEnabled)
{
	mDevices = new DevicePrefs [mMax];
	for (int i = 0; i < mNext; i++)
		mDevices[i] = other.mDevices[i];
}

CQSESPrefs& CQSESPrefs::operator=( const CQSESPrefs& source )
{

	if (&source != this)
	{
		mMax = source.mMax;
		mNext = source.mNext;
		mKeyCode = source.mKeyCode;
		mKeyMods = source.mKeyMods;
		mCycleKeyString = source.mCycleKeyString;
		mIsCycleKeyEnabled = source.mIsCycleKeyEnabled;

		delete[] mDevices;

		mDevices = new DevicePrefs[mMax];
		for (int i = 0; i < mNext; i++)
			mDevices[i] = source.mDevices[i];
	}

	return *this;
}

void CQSESPrefs::Add(DevicePrefs *sdi/*, bool def*/)
{
	int index = FindByID(sdi->DeviceID);

	if (index != -1)
	{
		mDevices[index] = *sdi;
		mDevices[index].IsPresent = true;
	}
	else
	{
		if (mNext == mMax)
			Remove(0);

		mDevices[mNext] = *sdi;
		mDevices[mNext].IsPresent = true;
		mNext++;
	}
}


void CQSESPrefs::Update(DevicePrefs *sdi/*, bool def*/)
{
	int index = FindByID(sdi->DeviceID);

	if (index != -1)
	{
		mDevices[index].DeviceID = sdi->DeviceID;
		mDevices[index].Name = sdi->Name;
		mDevices[index].IsPresent = true;
	}
	else
	{
		if (mNext == mMax)
			Remove(0);  //fixme

		mDevices[mNext] = *sdi;
		mDevices[mNext].IsPresent = true;
		mNext++;
	}
}

bool CQSESPrefs::SetMax(int max)
{
	if (max < mNext || max > mLimit)
		return false;
	mMax = max;
	return true;
}

wstring& CQSESPrefs::GetHotkeyString(int index, wstring& keyString)
{
	if (index < 0 || index >= mNext)
	{
		keyString.clear();
		return keyString;
	}

	keyString = mDevices[index].HotkeyString;
	return keyString;
}

void CQSESPrefs::SetHotkeyCode(int index, UINT code)
{
	if (index < 0 || index >= mNext)
		return;

	mDevices[index].KeyCode = code;

	if (mDevices[index].KeyCode == 0 && mDevices[index].KeyMods == 0)
		mDevices[index].HasHotkey = false;		mDevices[index].HasHotkey = false;
}

void CQSESPrefs::SetHotkeyMods(int index, UINT mods)
{
	if (index < mNext) mDevices[index].KeyMods = mods;

	if (mDevices[index].KeyCode == 0 && mDevices[index].KeyMods == 0)
		mDevices[index].HasHotkey = false;
}

UINT CQSESPrefs::GetHotkeyMods(int index)
{
	if (index < mNext)
		return mDevices[index].KeyMods;
	else return 0;
}


bool CQSESPrefs::GetHotkeyEnabled(int index)
{ 
	if (index < mNext)
		return mDevices[index].HasHotkey;
	return false;
}

bool CQSESPrefs::GetExcludeFromCycle(int index)
{ 
	if (index < mNext)
		return mDevices[index].IsExcludedFromCycle;
	return false;
}

bool CQSESPrefs::GetIsHidden(int index)
{ 
	if (index < mNext)
		return mDevices[index].IsHidden;
	return false;
}

bool CQSESPrefs::GetIsPresent(int index)
{ 
	if (index < mNext)
		return mDevices[index].IsPresent;
	return false;
}

void CQSESPrefs::Swap(int d1, int d2)
{ 
	DevicePrefs temp = mDevices[d1];
	mDevices[d1] = mDevices[d2];
	mDevices[d2] = temp;
}

void CQSESPrefs::Sort()
{
	for (int i = 0 ; i < mNext - 1; i++)
		if (!mDevices[i].IsPresent)
		{
			for(int j = i + 1; j < mNext; j++)
			{
				if (mDevices[j].IsPresent)
					Swap(i, j);
			}
		}
}


bool CQSESPrefs::GetByName(wstring& name, DevicePrefs * sdi)
{
	int index = FindByName(name);

	if(index != -1)
	{
		*sdi = mDevices[index];
		return true;
	}
	return false;
}

int CQSESPrefs::FindByID(const wstring& id)
{
	for (int i = 0; i < mNext; i++)
	{
		if (id == mDevices[i].DeviceID)
			return i;
	}
	return -1;
}

wstring& CQSESPrefs::GetName(int index, wstring& s)
{
	if (index < mNext)
		s = mDevices[index].Name;
	return s;
}

wstring& CQSESPrefs::GetID(int index, wstring& s)
{
	if (index < mNext)
		s = mDevices[index].DeviceID;
	return s;
}

bool CQSESPrefs::Init(int max)
{
	if (max < 1 || max > mLimit)
		return false;

	if (mDevices != 0)
		delete[] mDevices;

	mDevices = new DevicePrefs[max];

	if (mDevices == 0)
		return false;

	mMax = max;
	mNext = 0;

	for (int i = 0; i < mMax; i++)
		Erase(i);

	return true;
}

void CQSESPrefs::Remove(int index)
{
	if (index >= mNext)
		return;

	if (index < (mNext - 1))
	{
		for (int i = index; i < (mNext - 1); i++)
			mDevices[i] = mDevices[i + 1];
	}
	--mNext;
	Erase(mNext);
}

void CQSESPrefs::RemoveDupes()
{
	for (int i = 0 ; i < mNext - 1; i++)
		for (int j = i + 1; j < mNext; j++)
			if (mDevices[i].DeviceID == mDevices[j].DeviceID)
				Remove(j);
}

void CQSESPrefs::Clear(int index)
{
	mDevices[index].HotkeyString.clear();
	mDevices[index].KeyCode = 0;
	mDevices[index].KeyMods = 0;
	mDevices[index].HasHotkey = false;
	mDevices[index].IsExcludedFromCycle = false;
	mDevices[index].IsPresent = false;
	mDevices[index].IsHidden = false;
}

void CQSESPrefs::Erase(int index)
{
	mDevices[index].Name.clear();
	mDevices[index].DeviceID.clear();
	mDevices[index].HotkeyString.clear();
	mDevices[index].KeyCode = 0;
	mDevices[index].KeyMods = 0;
	mDevices[index].HasHotkey = false;
	mDevices[index].IsExcludedFromCycle = false;
	mDevices[index].IsPresent = false;
	mDevices[index].IsHidden = false;
}

void CQSESPrefs::ResetPresent()
{
	for (int i = 0; i < mMax; i++)
		mDevices[i].IsPresent = false;
}

int CQSESPrefs::FindByName(wstring& name)
{
	if (mNext != 0)
	{
		for (int i = 0; i < mNext; i++)
			if (name == mDevices[i].Name)
				return i;
	}
	return -1;
}

bool CQSESPrefs::Save()
{
	wstring FilePath;
	LPWSTR wstrPrefsPath = 0;

	RemoveDupes();

	SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, 0, &wstrPrefsPath);
	FilePath = wstrPrefsPath;
	CoTaskMemFree(wstrPrefsPath);

	FilePath += L"\\Audio Endpoint Switcher";
	CreateDirectoryW(FilePath.c_str(), 0);
	FilePath += L"\\Preferences.ini";

	FILE * fh;
	_wfopen_s(&fh, FilePath.c_str(), L"w");
	fwprintf(fh, L"[Globals]\n");
	fwprintf(fh, L"\tEnableCycling = %d\n", mIsCycleKeyEnabled);
	fwprintf(fh, L"\tKeyCode = 0x%x\n", mKeyCode);
	fwprintf(fh, L"\tKeyString = %s\n", mCycleKeyString.c_str());
	fwprintf(fh, L"\tAltKey = %d\n", mKeyMods & MOD_ALT);
	fwprintf(fh, L"\tControlKey = %d\n", mKeyMods & MOD_CONTROL);
	fwprintf(fh, L"\tShiftKey = %d\n", mKeyMods & MOD_SHIFT);
	fwprintf(fh, L"\tWinKey = %d\n", mKeyMods & MOD_WIN);

	for (int i = 0; i < mNext; i++)
	{
		fwprintf(fh, L"[%s]\n", mDevices[i].DeviceID.c_str());
		fwprintf(fh, L"\tName = %s\n", mDevices[i].Name.c_str());
		fwprintf(fh, L"\tHasKey = %d\n", mDevices[i].HasHotkey);
		fwprintf(fh, L"\tKeyString = %s\n", mDevices[i].HotkeyString.c_str());
		fwprintf(fh, L"\tExcluded = %d\n", mDevices[i].IsExcludedFromCycle);
		fwprintf(fh, L"\tHidden = %d\n", mDevices[i].IsHidden);
		fwprintf(fh, L"\tPresent = %d\n", mDevices[i].IsPresent);
		fwprintf(fh, L"\tKeyCode = 0x%x\n", mDevices[i].KeyCode);
		fwprintf(fh, L"\tAltKey = %d\n", mDevices[i].KeyMods & MOD_ALT);
		fwprintf(fh, L"\tControlKey = %d\n", mDevices[i].KeyMods & MOD_CONTROL);
		fwprintf(fh, L"\tShiftKey = %d\n", mDevices[i].KeyMods & MOD_SHIFT);
		fwprintf(fh, L"\tWinKey = %d\n", mDevices[i].KeyMods & MOD_WIN);
	}
	fclose(fh);
	
	return true;
}

bool CQSESPrefs::ReadConfig(const wstring &fileName, map <wstring, map <wstring, wstring> > & sections)
{
	const int cBufflen = 1024;

	wstring str, keyStr, valStr;
	wstring sectionStr(L"[]");
	size_t eq;
	wchar_t wBuffer[cBufflen];

	FILE * fh;
	if (_wfopen_s(&fh, fileName.c_str(), L"r") != 0)
		return false;

	while (fgetws(wBuffer, cBufflen, fh))
	{
		str = wBuffer;
		trim(str);
		if (str[0] == L'#')
			continue;
		size_t p = str.find(L'#');
		if (p != std::string::npos)
		{
			str.erase(p, str.length() - 1);
			rtrim(str);
		}
		if (str[0] == L'[')
		{
			transform(str.begin(), str.end(), str.begin(), ::tolower);
			sectionStr = str;
		}
		else if ((eq = str.find(L'=')) != string::npos)
		{
			keyStr = str.substr(0, eq);
			valStr = str.substr(eq + 1, str.length() - 1);
			rtrim(keyStr);
			ltrim(valStr);
			transform(keyStr.begin(), keyStr.end(), keyStr.begin(), ::tolower);
			sections[sectionStr][keyStr] = valStr;
		}
	}
	fclose(fh);
	return true;
}

bool CQSESPrefs::Load()
{
	LPWSTR wstrPrefsPath = 0;
	SHGetKnownFolderPath(FOLDERID_RoamingAppData, 0, 0, &wstrPrefsPath);
	wstring FilePath(wstrPrefsPath);
	CoTaskMemFree(wstrPrefsPath);

	FilePath += L"\\Audio Endpoint Switcher";
	CreateDirectoryW(FilePath.c_str(), 0);
	FilePath += L"\\Preferences.ini";

	map <wstring, map <wstring, wstring> > sections;

	// Fix: was missing 'return', silently ignoring load failures
	if (!ReadConfig(FilePath, sections))
		return false;

	if (sections.count(L"[globals]") != 0)
	{
		if (sections[L"[globals]"].count(L"keycode") != 0)
			mKeyCode = wcstoul(sections[L"[globals]"][L"keycode"].c_str(), nullptr, 16);
		if (sections[L"[globals]"].count(L"enablecycling") != 0)
			mIsCycleKeyEnabled = sections[L"[globals]"][L"enablecycling"] == L"1";
		if (sections[L"[globals]"].count(L"keystring") != 0)
			mCycleKeyString = sections[L"[globals]"][L"keystring"];
		if (sections[L"[globals]"].count(L"altkey") != 0 && sections[L"[globals]"][L"altkey"] == L"1")
			mKeyMods = MOD_ALT;
		if (sections[L"[globals]"].count(L"controlkey") != 0 && sections[L"[globals]"][L"controlkey"] == L"1")
			mKeyMods |= MOD_CONTROL;
		if (sections[L"[globals]"].count(L"shiftkey") != 0 && sections[L"[globals]"][L"shiftkey"] == L"1")
			mKeyMods |= MOD_SHIFT;
		if (sections[L"[globals]"].count(L"winkey") != 0 && sections[L"[globals]"][L"winkey"] == L"1")
			mKeyMods |= MOD_WIN;
	}

	wregex guid_regex(L"\\[\\{\\d\\.\\d\\.\\d\\.\\d{8}\\}\\.\\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\\}\\]");
	wsmatch wMatch;
	int i = 0;
	for (const auto& kv : sections)
	{
		if (!regex_match(kv.first.cbegin(), kv.first.cend(), wMatch, guid_regex))
			continue;
		if (i >= mMax)
			break;
		mDevices[i].DeviceID = kv.first.substr(1, kv.first.length() - 2);
		if (kv.second.count(L"name") != 0)
			mDevices[i].Name = kv.second.at(L"name");
		if (kv.second.count(L"altkey") != 0)
			mDevices[i].KeyMods = (kv.second.at(L"altkey") == L"1") ? MOD_ALT : 0;
		if (kv.second.count(L"controlkey") != 0 && kv.second.at(L"controlkey") == L"1")
			mDevices[i].KeyMods |= MOD_CONTROL;
		if (kv.second.count(L"shiftkey") != 0 && kv.second.at(L"shiftkey") == L"1")
			mDevices[i].KeyMods |= MOD_SHIFT;
		if (kv.second.count(L"winkey") != 0 && kv.second.at(L"winkey") == L"1")
			mDevices[i].KeyMods |= MOD_WIN;
		if (kv.second.count(L"keycode") != 0)
			mDevices[i].KeyCode = wcstoul(kv.second.at(L"keycode").c_str(), nullptr, 16);
		if (kv.second.count(L"keystring") != 0)
			mDevices[i].HotkeyString = kv.second.at(L"keystring");
		if (kv.second.count(L"haskey") != 0)
			mDevices[i].HasHotkey = kv.second.at(L"haskey") == L"1";
		if (kv.second.count(L"excluded") != 0)
			mDevices[i].IsExcludedFromCycle = kv.second.at(L"excluded") == L"1";
		if (kv.second.count(L"hidden") != 0)
			mDevices[i].IsHidden = kv.second.at(L"hidden") == L"1";
		if (kv.second.count(L"present") != 0)
			mDevices[i].IsPresent = kv.second.at(L"present") == L"1";

		i += 1;
	}

	mNext = i;
	return true;
}
