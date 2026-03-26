#pragma once

#include "mmdeviceapi.h"
#include <atlbase.h>
#include <atlcomcli.h>
#include <string>
#include "Prefs.h"

// User-defined window messages for audio endpoint notifications
inline constexpr UINT WM_USER_NOTIFICATION_DEFAULT = WM_USER + 2;
inline constexpr UINT WM_USER_NOTIFICATION_ADDED   = WM_USER + 3;
inline constexpr UINT WM_USER_NOTIFICATION_REMOVED = WM_USER + 4;
inline constexpr UINT WM_USER_NOTIFICATION_CHANGED = WM_USER + 5;

// CMMNotificationClient receives audio endpoint change notifications from the
// OS and forwards them as window messages.
//
// IUnknown is implemented manually with Interlocked ref-counting. This is
// intentional: the class is a single concrete type, never aggregated, and
// never needs a COM server or ATL module infrastructure.
class CMMNotificationClient : public IMMNotificationClient
{
public:
    explicit CMMNotificationClient(HWND hWnd = nullptr)
        : mhWnd(hWnd), mRefCount(1)
    {}

    void SetWindow(HWND hWnd) { mhWnd = hWnd; }

    // IUnknown
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return InterlockedIncrement(&mRefCount);
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        ULONG ref = InterlockedDecrement(&mRefCount);
        if (ref == 0)
            delete this;
        return ref;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvInterface) override
    {
        if (riid == IID_IUnknown)
        {
            AddRef();
            *ppvInterface = static_cast<IUnknown*>(this);
        }
        else if (riid == __uuidof(IMMNotificationClient))
        {
            AddRef();
            *ppvInterface = static_cast<IMMNotificationClient*>(this);
        }
        else
        {
            *ppvInterface = nullptr;
            return E_NOINTERFACE;
        }
        return S_OK;
    }

    // IMMNotificationClient
    HRESULT STDMETHODCALLTYPE OnDefaultDeviceChanged(
        EDataFlow /*flow*/, ERole role, LPCWSTR /*pwstrDeviceId*/) override
    {
        if (role == eConsole)
            SendMessage(mhWnd, WM_USER_NOTIFICATION_DEFAULT, 0, 0);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnDeviceAdded(LPCWSTR /*pwstrDeviceId*/) override
    {
        SendMessage(mhWnd, WM_USER_NOTIFICATION_ADDED, 0, 0);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnDeviceRemoved(LPCWSTR /*pwstrDeviceId*/) override
    {
        SendMessage(mhWnd, WM_USER_NOTIFICATION_REMOVED, 0, 0);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnDeviceStateChanged(
        LPCWSTR /*pwstrDeviceId*/, DWORD /*dwNewState*/) override
    {
        SendMessage(mhWnd, WM_USER_NOTIFICATION_CHANGED, 0, 0);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE OnPropertyValueChanged(
        LPCWSTR /*pwstrDeviceId*/, const PROPERTYKEY /*key*/) override
    {
        return S_OK;
    }

private:
    HWND mhWnd   = nullptr;
    LONG mRefCount = 1;
};
