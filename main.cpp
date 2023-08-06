// SPDX-License-Identifier: GPL-3.0

#include <array>
#include <cstdint>
#include <limits>
#include <memory>
#include <stdexcept>
#include <type_traits>
#include <utility>

#include <Richedit.h>
#include <Windows.h>
#include <endpointvolume.h>
#include <mmdeviceapi.h>
#include <objbase.h>
#include <shellapi.h>

#ifndef NDEBUG
#  define GET_ERROR() \
    do { \
      lastError = GetLastError(); \
    } while (false)
static DWORD lastError;
#else
#  define GET_ERROR()
#endif

#define TRY_HR(x) \
  do { \
    if (auto _result = (x); FAILED(_result)) { \
      GET_ERROR(); \
      throw std::runtime_error(#x " failed"); \
    } \
  } while (false)

#define TRY_BOOL(x) \
  do { \
    if (auto _result = (x); _result == 0) { \
      GET_ERROR(); \
      throw std::runtime_error(#x " failed"); \
    } \
  } while (false)

#define PRECONDITION(x) \
  do { \
    if (!(x)) { \
      throw std::runtime_error("Precondition [ " #x " ] not met"); \
    } \
  } while (false)

#define THROW_IF(condition, message) \
  do { \
    if (condition) { \
      GET_ERROR(); \
      throw std::runtime_error(message); \
    } \
  } while (false)

namespace
{

struct UserMessage
{
  enum : UINT
  {
    TrayIcon = WM_USER,
    GetDefaultEndpoint = WM_USER + 1,
    ChangeAudio = WM_USER + 2,
  };
  static constexpr WORD TrayLicense = 1;
  static constexpr WORD TrayExit = 2;
};

template<typename Derived, typename Base>
concept DerivedFrom = std::is_base_of_v<Base, Derived>;

using ComPtrDeleter = decltype([](IUnknown* ptr) { if (ptr != nullptr) ptr->Release(); });

template<DerivedFrom<IUnknown> T>
using ComPtr = std::unique_ptr<T, ComPtrDeleter>;

struct Handle
{
  HANDLE handle {};

  Handle(HANDLE handle)
      : handle(handle)
  {
  }

  ~Handle() noexcept(false)
  {
    THROW_IF(handle != nullptr && CloseHandle(handle) == 0,
             "CloseHandle failed");
  }
};

struct Library
{
  HINSTANCE library {};

  Library(LPCWSTR name)
      : library(LoadLibraryW(name))
  {
    THROW_IF(library == nullptr, "LoadLibrary failed");
  }

  ~Library() noexcept(false)
  {
    THROW_IF(library != nullptr && FreeLibrary(library) == 0,
             "FreeLibrary failed");
  }
};

class TrayIcon
{
  NOTIFYICONDATAW* iconData {};

public:
  TrayIcon(NOTIFYICONDATAW& iconData)
      : iconData(&iconData)
  {
    THROW_IF(Shell_NotifyIconW(NIM_ADD, &iconData) == FALSE,
             "Shell_NotifyIconW(NIM_ADD) failed");
  }

  ~TrayIcon() noexcept(false)
  {
    THROW_IF(Shell_NotifyIconW(NIM_DELETE, iconData) == FALSE,
             "Shell_NotifyIconW(NIM_DELETE) failed");
  }
};

class EndpointHandler : public IAudioEndpointVolumeCallback
{
  ULONG refCount {};
  GUID* guid {};
  HWND window {};

public:
  EndpointHandler(GUID* guid, HWND window)
      : guid(guid)
      , window(window)
  {
  }

  EndpointHandler(EndpointHandler&&) = delete;

  virtual ~EndpointHandler() {}

  ULONG STDMETHODCALLTYPE AddRef() override
  {
    PRECONDITION(refCount != std::numeric_limits<ULONG>::max());
    return InterlockedIncrement(&refCount);
  }

  ULONG STDMETHODCALLTYPE Release() override
  {
    PRECONDITION(refCount > 0);
    auto result = InterlockedDecrement(&refCount);
    if (result == 0) {
      delete this;
    }
    return result;
  }

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,
                                           VOID** ppvInterface) override
  {
    if (IID_IUnknown != riid || __uuidof(IAudioEndpointVolumeCallback) != riid)
    {
      *ppvInterface = nullptr;
      return E_NOINTERFACE;
    }

    AddRef();
    *ppvInterface = this;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE
  OnNotify(PAUDIO_VOLUME_NOTIFICATION_DATA data) override
  {
    if (data == nullptr) {
      return E_POINTER;
    }

    if (data->guidEventContext != *guid) {
      if (PostMessageW(window, UserMessage::ChangeAudio, 0, 0) == 0) {
        auto error = GetLastError();
        return __HRESULT_FROM_WIN32(error);
      }
    }

    return S_OK;
  }
};

class NotificationClient : public IMMNotificationClient
{
  ULONG refCount {};
  HWND window {};

public:
  NotificationClient(HWND window)
      : window(window)
  {
  }

  NotificationClient(NotificationClient&&) = delete;

  virtual ~NotificationClient() {}

  ULONG STDMETHODCALLTYPE AddRef() override
  {
    PRECONDITION(refCount != std::numeric_limits<ULONG>::max());
    return InterlockedIncrement(&refCount);
  }

  ULONG STDMETHODCALLTYPE Release() override
  {
    PRECONDITION(refCount > 0);
    auto result = InterlockedDecrement(&refCount);
    if (result == 0) {
      delete this;
    }
    return result;
  }

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,
                                           VOID** ppvInterface) override
  {
    if (IID_IUnknown != riid || __uuidof(IMMNotificationClient) != riid) {
      *ppvInterface = nullptr;
      return E_NOINTERFACE;
    }

    AddRef();
    *ppvInterface = this;
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE OnDefaultDeviceChanged(EDataFlow flow,
                                                   ERole role,
                                                   LPCWSTR) override
  {
    if (flow != eRender || role != eConsole) {
      return S_OK;
    }

    if (PostMessageW(window, UserMessage::GetDefaultEndpoint, 0, 0) == 0) {
      auto error = GetLastError();
      return __HRESULT_FROM_WIN32(error);
    }

    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE OnDeviceStateChanged(LPCWSTR, DWORD) override
  {
    return S_OK;
  }

  HRESULT STDMETHODCALLTYPE OnDeviceAdded(LPCWSTR) override { return S_OK; }

  HRESULT STDMETHODCALLTYPE OnDeviceRemoved(LPCWSTR) override { return S_OK; }

  HRESULT STDMETHODCALLTYPE OnPropertyValueChanged(  //
      LPCWSTR,
      const PROPERTYKEY) override
  {
    return S_OK;
  }
};

void ShowContextMenu(HWND hwnd)
{
  auto popup = CreatePopupMenu();
  THROW_IF(popup == nullptr, "CreatePopupMenu failed");

  TRY_BOOL(InsertMenuW(popup,
                       0,
                       MF_BYPOSITION | MF_STRING,
                       UserMessage::TrayLicense,
                       L"License"));

  TRY_BOOL(InsertMenuW(
      popup, 1, MF_BYPOSITION | MF_STRING, UserMessage::TrayExit, L"Exit"));

  (void)SetForegroundWindow(hwnd);

  auto point = POINT {};
  TRY_BOOL(GetCursorPos(&point));

  TRY_BOOL(TrackPopupMenu(popup, 0, point.x, point.y, 0, hwnd, nullptr));
}

class LicenseDialogData
{
public:
  static constexpr WORD buttonId = 1;
  static constexpr WORD textId = 2;

private:
  template<std::size_t Size>
  struct Storage
  {
    static_assert(Size != 0);
    alignas(DLGTEMPLATE) std::byte buffer[Size] {};
  };

  static constexpr auto precalculatedDialogData = []
  {
    constexpr auto oversized = []
    {
      auto buffer = std::array<std::byte, 2048> {};
      auto usedBytes = 0ULL;
      auto appendData = [&]<typename T>(T const& data) -> void
      {
        auto bytes = std::bit_cast<std::array<std::byte, sizeof(T)>>(data);
        std::copy(bytes.begin(), bytes.end(), buffer.data() + usedBytes);
        usedBytes += sizeof(T);
      };
      auto alignBuffer = [&](std::size_t alignment) -> void
      { usedBytes += usedBytes % alignment; };

      appendData(DLGTEMPLATE {
          .style = DS_SETFONT | DS_MODALFRAME | DS_FIXEDSYS | WS_POPUP
              | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
          .dwExtendedStyle = 0,
          .cdit = 2,
          .x = 0,
          .y = 0,
          .cx = 249,
          .cy = 152,
      });
      appendData(WORD {0});
      appendData(WORD {0});
      appendData(L"License");
      appendData(WORD {8});
      appendData(L"MS Shell Dlg");

      alignBuffer(sizeof(DWORD));
      appendData(DLGITEMTEMPLATE {
          .style = BS_DEFPUSHBUTTON | WS_VISIBLE,
          .dwExtendedStyle = 0,
          .x = 192,
          .y = 131,
          .cx = 50,
          .cy = 14,
          .id = buttonId,
      });
      appendData(WORD {0xFFFF});
      appendData(WORD {0x0080});
      appendData(L"OK");
      appendData(WORD {0});

      alignBuffer(sizeof(DWORD));
      appendData(DLGITEMTEMPLATE {
          .style = ES_MULTILINE | ES_READONLY | WS_VISIBLE,
          .dwExtendedStyle = 0,
          .x = 7,
          .y = 7,
          .cx = 235,
          .cy = 118,
          .id = textId,
      });
      appendData(RICHEDIT_CLASSW);
      appendData(WORD {0});
      appendData(WORD {0});

      return std::make_pair(buffer, usedBytes);
    }();

    auto storage = Storage<oversized.second> {};
    std::copy(oversized.first.data(),
              oversized.first.data() + oversized.second,
              storage.buffer);
    return storage;
  }();

public:
  static LPCDLGTEMPLATEW get()
  {
    return reinterpret_cast<LPCDLGTEMPLATEW>(precalculatedDialogData.buffer);
  }
};

struct State
{
  HINSTANCE hInstance {};
  HWND dialog {};
  GUID* guid {};
  IMMDeviceEnumerator* deviceEnumerator {};
  ComPtr<IMMDevice> audioDevice {};
  ComPtr<IAudioEndpointVolume> endpointVolume {};
};

wchar_t const* gplNotice =
    L""
    "AlwaysMute to keep the default audio device on Windows quiet\n"
    "Copyright (C) 2023 friendlyanon\n\n"
    "AlwaysMute is free software: you can redistribute it and/or modify\n"
    "it under the terms of the GNU General Public License as published by\n"
    "the Free Software Foundation, version 3.\n\n"
    "AlwaysMute is distributed in the hope that it will be useful,\n"
    "but WITHOUT ANY WARRANTY; without even the implied warranty of\n"
    "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the\n"
    "GNU General Public License for more details.\n\n"
    "You should have received a copy of the GNU General Public License\n"
    "along with AlwaysMute. If not, see <https://www.gnu.org/licenses/>.";

INT_PTR CALLBACK DialogProc(  //
    HWND hwnd,
    UINT message,
    WPARAM wParam,
    LPARAM lParam)
{
  switch (message) {
    case WM_INITDIALOG: {
      SetLastError(0);
      if (SetWindowLongPtrW(hwnd, GWLP_USERDATA, lParam) == 0) {
        THROW_IF(GetLastError() != 0, "SetWindowLongPtrW failed");
      }

      (void)SendDlgItemMessageW(
          hwnd, LicenseDialogData::textId, EM_SETEVENTMASK, 0, ENM_LINK);

      THROW_IF(SendDlgItemMessageW(  //
                   hwnd,
                   LicenseDialogData::textId,
                   EM_AUTOURLDETECT,
                   AURL_ENABLEURL,
                   0)
                   != 0,
               "Couldn't set URL autodetection");

      THROW_IF(SendDlgItemMessageW(hwnd,
                                   LicenseDialogData::textId,
                                   WM_SETTEXT,
                                   0,
                                   reinterpret_cast<LPARAM>(gplNotice))
                   != TRUE,
               "Couldn't set GPL notice");

      return TRUE;
    }
    case WM_NOTIFY:
      switch (reinterpret_cast<LPNMHDR>(lParam)->code) {
        case EN_LINK: {
          if (LOWORD(wParam) != LicenseDialogData::textId) {
            break;
          }
          auto* enlink = reinterpret_cast<ENLINK*>(lParam);
          switch (enlink->msg) {
            case WM_LBUTTONUP: {
              auto buffer = std::make_unique<wchar_t[]>(
                  static_cast<std::size_t>(enlink->chrg.cpMax)
                  - enlink->chrg.cpMin + 1);
              auto range = TEXTRANGE {enlink->chrg, buffer.get()};
              (void)SendDlgItemMessageW(  //
                  hwnd,
                  LicenseDialogData::textId,
                  EM_GETTEXTRANGE,
                  0,
                  reinterpret_cast<LPARAM>(&range));
              (void)ShellExecuteW(
                  hwnd, L"open", buffer.get(), nullptr, nullptr, SW_SHOW);
              return TRUE;
            }
          }
          break;
        }
      }
      break;
    case WM_COMMAND:
      switch (LOWORD(wParam)) {
        case LicenseDialogData::buttonId:
          TRY_BOOL(DestroyWindow(hwnd));
          return TRUE;
      }
      break;
    case WM_CLOSE:
      TRY_BOOL(DestroyWindow(hwnd));
      return TRUE;
    case WM_DESTROY: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      THROW_IF(userData == 0, "GetWindowLongPtrW failed");
      *reinterpret_cast<HWND*>(userData) = nullptr;
      return TRUE;
    }
  }

  return FALSE;
}

void ChangeAudio(State& state)
{
  TRY_HR(state.endpointVolume->SetMasterVolumeLevelScalar(0.0f, state.guid));
}

LRESULT CALLBACK MainWndProc(  //
    HWND hwnd,
    UINT message,
    WPARAM wParam,
    LPARAM lParam)
{
  switch (message) {
    case WM_CREATE:
      SetLastError(0);
      if (SetWindowLongPtrW(
              hwnd,
              GWLP_USERDATA,
              reinterpret_cast<LONG_PTR>(
                  reinterpret_cast<CREATESTRUCT*>(lParam)->lpCreateParams))
          != 0)
      {
        break;
      }
      THROW_IF(GetLastError() != 0, "SetWindowLongPtrW failed");
      break;
    case UserMessage::TrayIcon:
      switch (LOWORD(lParam)) {
        case WM_RBUTTONDOWN:
          ShowContextMenu(hwnd);
          break;
      }
      break;
    case UserMessage::GetDefaultEndpoint: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      THROW_IF(userData == 0, "GetWindowLongPtrW failed");

      auto* state = reinterpret_cast<State*>(userData);
      state->endpointVolume.reset();
      state->audioDevice.reset();

      if (auto result = state->deviceEnumerator->GetDefaultAudioEndpoint(
              eRender, eConsole, std::out_ptr(state->audioDevice));
          result == E_NOTFOUND)
      {
        break;
      } else {
        THROW_IF(FAILED(result), "GetDefaultAudioEndpoint failed");
      }

      TRY_HR(state->audioDevice->Activate(__uuidof(IAudioEndpointVolume),
                                          CLSCTX_INPROC_SERVER,
                                          nullptr,
                                          std::out_ptr(state->endpointVolume)));
      TRY_HR(state->endpointVolume->RegisterControlChangeNotify(
          new EndpointHandler(state->guid, hwnd)));
      ChangeAudio(*state);
      break;
    }
    case UserMessage::ChangeAudio: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      THROW_IF(userData == 0, "GetWindowLongPtrW failed");

      ChangeAudio(*reinterpret_cast<State*>(userData));
      break;
    }
    case WM_COMMAND:
      switch (LOWORD(wParam)) {
        case UserMessage::TrayLicense: {
          auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
          THROW_IF(userData == 0, "GetWindowLongPtrW failed");

          auto& state = *reinterpret_cast<State*>(userData);
          if (state.dialog != nullptr) {
            break;
          }

          state.dialog = CreateDialogIndirectParamW(  //
              state.hInstance,
              LicenseDialogData::get(),
              nullptr,
              DialogProc,
              reinterpret_cast<LPARAM>(&state.dialog));
          THROW_IF(state.dialog == nullptr, "Couldn't create license dialog");
          break;
        }
        case UserMessage::TrayExit:
          TRY_BOOL(DestroyWindow(hwnd));
          break;
      }
      break;
    case WM_CLOSE:
      TRY_BOOL(DestroyWindow(hwnd));
      break;
    case WM_DESTROY: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      THROW_IF(userData == 0, "GetWindowLongPtrW failed");

      auto* state = reinterpret_cast<State*>(userData);
      if (state->dialog != nullptr) {
        TRY_BOOL(DestroyWindow(state->dialog));
      }

      PostQuitMessage(0);
      break;
    }
  }

  return DefWindowProcW(hwnd, message, wParam, lParam);
}

int TryMain(HINSTANCE hInstance)
{
  auto mutex = Handle(CreateMutexW(nullptr, FALSE, L"Local\\AlwaysMute"));
  if (GetLastError() == ERROR_ALREADY_EXISTS) {
    return 0;
  }

  TRY_HR(CoInitialize(nullptr));

  auto guid = GUID {};
  TRY_HR(CoCreateGuid(&guid));

  auto deviceEnumerator = ComPtr<IMMDeviceEnumerator>();
  TRY_HR(CoCreateInstance(__uuidof(MMDeviceEnumerator),
                          nullptr,
                          CLSCTX_INPROC_SERVER,
                          __uuidof(IMMDeviceEnumerator),
                          std::out_ptr(deviceEnumerator)));

  auto richEdit = Library(L"Riched20.dll");
  auto cursor = LoadCursorW(nullptr, IDC_ARROW);
  THROW_IF(cursor == nullptr, "LoadCursorW failed");

  auto mainWindowClass = WNDCLASSW  //
      {.lpfnWndProc = MainWndProc,
       .hInstance = hInstance,
       .hCursor = cursor,
       .lpszClassName = L"AlwaysMute - Main"};
  auto mainAtom = RegisterClassW(&mainWindowClass);
  THROW_IF(mainAtom == 0, "RegisterClassW failed");

  auto state = State  //
      {.hInstance = hInstance,
       .guid = &guid,
       .deviceEnumerator = deviceEnumerator.get()};
  auto window = CreateWindowW(  //
      MAKEINTATOM(mainAtom),
      L"Message only",
      0,
      0,
      0,
      0,
      0,
      HWND_MESSAGE,
      nullptr,
      hInstance,
      &state);
  THROW_IF(window == nullptr, "CreateWindowW failed");

  TRY_HR(deviceEnumerator->RegisterEndpointNotificationCallback(
      new NotificationClient(window)));

  auto sndVol = Library(L"SndVolSSO.dll");
  auto icon = LoadIconW(sndVol.library, MAKEINTRESOURCEW(120));
  THROW_IF(icon == nullptr, "LoadIconW failed");

  auto trayIconData = NOTIFYICONDATAW {
      .cbSize = sizeof(NOTIFYICONDATAW),
      .hWnd = window,
      .uID = 0,
      .uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP,
      .uCallbackMessage = UserMessage::TrayIcon,
      .hIcon = icon,
      .szTip = L"AlwaysMute",
  };
  auto trayIcon = TrayIcon(trayIconData);

  auto msg = MSG {};
  (void)PeekMessageW(&msg, window, 0, 0, PM_NOREMOVE);
  TRY_BOOL(PostMessageW(window, UserMessage::GetDefaultEndpoint, 0, 0));

  while (true) {
    if (auto result = GetMessageW(&msg, nullptr, 0, 0); result == 0) {
      break;
    } else {
      THROW_IF(result == -1, "GetMessageW failed");
    }

    if (state.dialog != nullptr && IsDialogMessageW(state.dialog, &msg) != 0) {
      continue;
    }

    (void)DispatchMessageW(&msg);
  }

  return static_cast<int>(msg.wParam);
}

}  // namespace

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int)
{
  try {
    return TryMain(hInstance);
  } catch (std::exception const& error) {
    OutputDebugStringA(error.what());
    OutputDebugStringA("\n");

    return 1;
  }
}