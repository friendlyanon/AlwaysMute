// SPDX-License-Identifier: GPL-3.0

#include <algorithm>
#include <array>
#include <bit>
#include <cstdint>
#include <exception>
#include <iostream>
#include <iterator>
#include <limits>
#include <memory>
#include <new>
#include <ranges>
#include <stacktrace>
#include <stdexcept>
#include <string_view>
#include <system_error>
#include <type_traits>
#include <utility>

#include <Richedit.h>
#include <Windows.h>
#include <endpointvolume.h>
#include <mmdeviceapi.h>
#include <objbase.h>
#include <shellapi.h>

#define ADDREF_RELEASE_METHODS \
  ULONG STDMETHODCALLTYPE AddRef() override \
  { \
    auto result = InterlockedIncrement(&refCount); \
    if (static_cast<LONG>(result) < 0) { \
      std::cerr << std::stacktrace::current() << '\n'; \
      throw std::runtime_error("refCount went beyond LONG_MAX"); \
    } \
    return result; \
  } \
\
  ULONG STDMETHODCALLTYPE Release() override \
  { \
    auto result = InterlockedDecrement(&refCount); \
    if (static_cast<LONG>(result) < 0) { \
      std::cerr << std::stacktrace::current() << '\n'; \
      throw std::runtime_error("refCount was not bigger than 0"); \
    } \
    if (result == 0) { \
      delete this; \
    } \
    return result; \
  }

using namespace std::string_view_literals;

namespace
{

class com_error : public std::exception
{
  HRESULT _code;

public:
  explicit com_error(HRESULT code)
      : _code(code)
  {
  }

  HRESULT code() const { return _code; }
};

void outputSystemError(DWORD error = GetLastError())
{
  constexpr auto bufferSize = DWORD {4096};
  auto buffer = std::array<wchar_t, bufferSize>();
  auto charactersWrittenWithoutNull = FormatMessageW(  //
      FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
      nullptr,
      error,
      MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL),
      buffer.data(),
      bufferSize,
      nullptr);
  if (charactersWrittenWithoutNull != 0) {
    auto it = buffer.begin() + charactersWrittenWithoutNull;
    auto tail = L"\n\0"sv;
    if (auto distance = static_cast<std::size_t>(buffer.end() - it);
        tail.size() > distance)
    {
      tail.remove_prefix(tail.size() - distance);
    }
    (void)std::ranges::copy(tail, it);
    OutputDebugStringW(buffer.data());
  } else {
    OutputDebugStringW(L"Can't get error message\n");
  }
}

[[noreturn]] void throw_(
    DWORD error, std::stacktrace stacktrace = std::stacktrace::current())
{
  std::cerr << stacktrace << '\n';
  throw std::system_error(static_cast<int>(error), std::system_category());
}

void throwIf(bool condition,
             DWORD error,
             std::stacktrace stacktrace = std::stacktrace::current())
{
  if (condition) {
    throw_(error, stacktrace);
  }
}

void throwIf(bool condition,
             std::stacktrace stacktrace = std::stacktrace::current())
{
  throwIf(condition, GetLastError(), stacktrace);
}

void throwIfCOM(HRESULT result,
                std::stacktrace stacktrace = std::stacktrace::current())
{
  if (FAILED(result)) {
    std::cerr << stacktrace << '\n';
    throw com_error(result);
  }
}

template<typename T>
T* as_ptr(auto value)
{
  return std::launder(reinterpret_cast<T*>(value));
}

struct UserMessage
{
  enum enum_ : UINT
  {
    TrayIcon = WM_USER,
    GetDefaultEndpoint,
    ChangeAudio,
  };
  static constexpr WORD TrayLicense = 1;
  static constexpr WORD TrayExit = 2;
};

template<typename Derived, typename Base>
concept DerivedFrom = std::is_base_of_v<Base, Derived>;

struct ComPtrDeleter
{
  void operator()(IUnknown* ptr)
  {
    if (ptr != nullptr) {
      ptr->Release();
    }
  }
};

template<DerivedFrom<IUnknown> T>
using ComPtr = std::unique_ptr<T, ComPtrDeleter>;

struct Handle
{
  HANDLE handle {};

  explicit Handle(HANDLE handle)
      : handle(handle)
  {
  }

  Handle(Handle&&) = delete;

  ~Handle()
  {
    if (handle != nullptr && CloseHandle(handle) == 0) {
      std::cerr << std::stacktrace::current() << '\n';
      outputSystemError();
    }
  }
};

struct Library
{
  HINSTANCE library {};

  explicit Library(LPCWSTR name)
      : library(LoadLibraryW(name))
  {
    throwIf(library == nullptr);
  }

  Library(Library&&) = delete;

  ~Library()
  {
    if (library != nullptr && FreeLibrary(library) == 0) {
      std::cerr << std::stacktrace::current() << '\n';
      outputSystemError();
    }
  }
};

class TrayIcon
{
  NOTIFYICONDATAW* iconData {};

public:
  explicit TrayIcon(NOTIFYICONDATAW& iconData)
      : iconData(&iconData)
  {
    throwIf(Shell_NotifyIconW(NIM_ADD, &iconData) == FALSE);
  }

  TrayIcon(TrayIcon&&) = delete;

  ~TrayIcon()
  {
    if (Shell_NotifyIconW(NIM_DELETE, iconData) == FALSE) {
      std::cerr << std::stacktrace::current() << '\n';
      OutputDebugStringW(L"Shell_NotifyIconW(NIM_DELETE) failed\n");
    }
  }
};

template<DerivedFrom<IUnknown> T>
class ComCallback
{
  T* callback;

public:
  template<typename... Args>
  explicit ComCallback(Args&&... args)
      : callback(new T(std::forward<Args>(args)...))
  {
  }

  ~ComCallback()
  {
    if (std::uncaught_exceptions() != 0) {
      callback->Release();
    }
  }

  operator T*() { return callback; }
};

class EndpointHandler : public IAudioEndpointVolumeCallback
{
  ULONG refCount {};
  GUID* guid {};
  HWND window {};

  virtual ~EndpointHandler() {}

public:
  EndpointHandler(GUID& guid, HWND window)
      : guid(&guid)
      , window(window)
  {
  }

  EndpointHandler(EndpointHandler&&) = delete;

  ADDREF_RELEASE_METHODS

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,
                                           void** ppvObject) override
  {
    if (ppvObject == nullptr) {
      return E_POINTER;
    }

    if (__uuidof(IUnknown) != riid
        || __uuidof(IAudioEndpointVolumeCallback) != riid)
    {
      *ppvObject = nullptr;
      return E_NOINTERFACE;
    }

    AddRef();
    *ppvObject = this;
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

  virtual ~NotificationClient() {}

public:
  explicit NotificationClient(HWND window)
      : window(window)
  {
  }

  NotificationClient(NotificationClient&&) = delete;

  ADDREF_RELEASE_METHODS

  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,
                                           void** ppvObject) override
  {
    if (ppvObject == nullptr) {
      return E_POINTER;
    }

    if (__uuidof(IUnknown) != riid || __uuidof(IMMNotificationClient) != riid) {
      *ppvObject = nullptr;
      return E_NOINTERFACE;
    }

    AddRef();
    *ppvObject = this;
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

struct PopupMenu
{
  HMENU popup {};

  PopupMenu()
      : popup(CreatePopupMenu())
  {
    throwIf(popup == nullptr);
  }

  PopupMenu(PopupMenu&&) = delete;

  ~PopupMenu()
  {
    if (popup != nullptr && DestroyMenu(popup) == 0) {
      std::cerr << std::stacktrace::current() << '\n';
      outputSystemError();
    }
  }
};

void ShowContextMenu(HWND hwnd)
{
  auto popup = PopupMenu();

  throwIf(InsertMenuW(popup.popup,
                      0,
                      MF_BYPOSITION | MF_STRING,
                      UserMessage::TrayLicense,
                      L"License")
          == 0);

  throwIf(InsertMenuW(popup.popup,
                      1,
                      MF_BYPOSITION | MF_STRING,
                      UserMessage::TrayExit,
                      L"Exit")
          == 0);

  (void)SetForegroundWindow(hwnd);

  auto point = POINT {};
  throwIf(GetCursorPos(&point) == 0);

  throwIf(TrackPopupMenu(popup.popup, 0, point.x, point.y, 0, hwnd, nullptr)
          == 0);
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

    // DLGTEMPLATE is a packed struct and its largest members are DWORD (4)
    alignas(4) std::byte buffer[Size] {};
  };

  static constexpr auto precalculatedDialogData = []
  {
    constexpr auto oversized = []
    {
      using u16 = std::uint16_t;

      auto buffer = std::array<std::byte, 256> {};
      auto usedBytes = 0ULL;
      auto data = [&]<typename T>(T const& data_) -> void
      {
        constexpr auto size = sizeof(T);
        if (size + usedBytes >= buffer.size()) {
          throw std::runtime_error("Buffer not big enough");
        }

        auto bytes = std::bit_cast<std::array<std::byte, size>>(data_);
        std::copy(bytes.begin(), bytes.end(), buffer.data() + usedBytes);
        usedBytes += size;
      };
      auto align = [&](std::size_t alignment) -> void
      { usedBytes += usedBytes % alignment; };
      auto trail = [&](auto const& data_) -> void
      {
        align(sizeof(WORD));
        data(data_);
      };
      auto item = [&](auto const& data_) -> void
      {
        align(sizeof(DWORD));
        data(data_);
      };

      data(DLGTEMPLATE {
          .style = DS_SETFONT | DS_MODALFRAME | DS_FIXEDSYS | WS_POPUP
              | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
          .dwExtendedStyle = 0,
          .cdit = 2,
          .x = 0,
          .y = 0,
          .cx = 249,
          .cy = 152,
      });
      trail(u16 {0});
      trail(u16 {0});
      trail(L"License");
      trail(u16 {8});
      trail(L"MS Shell Dlg");

      item(DLGITEMTEMPLATE {
          .style = BS_DEFPUSHBUTTON | WS_VISIBLE,
          .dwExtendedStyle = 0,
          .x = 192,
          .y = 131,
          .cx = 50,
          .cy = 14,
          .id = buttonId,
      });
      trail(u16 {0xFFFF});
      trail(u16 {0x0080});
      trail(L"OK");
      trail(u16 {0});

      item(DLGITEMTEMPLATE {
          .style = ES_MULTILINE | ES_READONLY | WS_VISIBLE,
          .dwExtendedStyle = 0,
          .x = 7,
          .y = 7,
          .cx = 235,
          .cy = 118,
          .id = textId,
      });
      trail(RICHEDIT_CLASSW);
      trail(u16 {0});
      trail(u16 {0});

      return std::make_pair(buffer, usedBytes);
    }();

    constexpr auto size = oversized.second;
    auto storage = Storage<size> {};
    auto it = oversized.first.begin();
    std::copy(it, it + size, storage.buffer);
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

#define GPL_URL L"https://www.gnu.org/licenses/"

wchar_t const* gplNotice =
    L""
    "AlwaysMute to keep the default audio device on Windows quiet\n"
    "Copyright (C) 2025 friendlyanon\n\n"
    "AlwaysMute is free software: you can redistribute it and/or modify\n"
    "it under the terms of the GNU General Public License as published by\n"
    "the Free Software Foundation, version 3.\n\n"
    "AlwaysMute is distributed in the hope that it will be useful,\n"
    "but WITHOUT ANY WARRANTY; without even the implied warranty of\n"
    "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the\n"
    "GNU General Public License for more details.\n\n"
    "You should have received a copy of the GNU General Public License\n"
    "along with AlwaysMute. If not, see <" GPL_URL ">.";

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
        auto error = GetLastError();
        throwIf(error != 0, error);
      }

      (void)SendDlgItemMessageW(
          hwnd, LicenseDialogData::textId, EM_SETEVENTMASK, 0, ENM_LINK);

      throwIf(SendDlgItemMessageW(  //
                  hwnd,
                  LicenseDialogData::textId,
                  EM_AUTOURLDETECT,
                  AURL_ENABLEURL,
                  0)
              != 0);

      throwIf(SendDlgItemMessageW(  //
                  hwnd,
                  LicenseDialogData::textId,
                  WM_SETTEXT,
                  0,
                  reinterpret_cast<LPARAM>(gplNotice))
              != TRUE);

      return TRUE;
    }
    case WM_NOTIFY:
      switch (as_ptr<NMHDR>(lParam)->code) {
        case EN_LINK: {
          if (LOWORD(wParam) != LicenseDialogData::textId) {
            break;
          }
          switch (as_ptr<ENLINK>(lParam)->msg) {
            case WM_LBUTTONUP:
              (void)ShellExecuteW(
                  hwnd, L"open", GPL_URL, nullptr, nullptr, SW_SHOW);
              return TRUE;
          }
          break;
        }
      }
      break;
    case WM_COMMAND:
      switch (LOWORD(wParam)) {
        case LicenseDialogData::buttonId:
          throwIf(DestroyWindow(hwnd) == 0);
          return TRUE;
      }
      break;
    case WM_CLOSE:
      throwIf(DestroyWindow(hwnd) == 0);
      return TRUE;
    case WM_DESTROY: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      throwIf(userData == 0);
      *as_ptr<HWND>(userData) = nullptr;
      return TRUE;
    }
  }

  return FALSE;
}

void ChangeAudio(State& state)
{
  throwIfCOM(
      state.endpointVolume->SetMasterVolumeLevelScalar(0.0f, state.guid));
}

LRESULT CALLBACK MainWndProc(  //
    HWND hwnd,
    UINT message,
    WPARAM wParam,
    LPARAM lParam)
{
  switch (message) {
    case WM_CREATE: {
      auto* state = as_ptr<CREATESTRUCT>(lParam)->lpCreateParams;
      SetLastError(0);
      if (SetWindowLongPtrW(
              hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state))
          != 0)
      {
        break;
      }
      throwIf(GetLastError() != 0);
      break;
    }
    case UserMessage::TrayIcon:
      switch (LOWORD(lParam)) {
        case WM_RBUTTONDOWN:
          ShowContextMenu(hwnd);
          break;
      }
      break;
    case UserMessage::GetDefaultEndpoint: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      throwIf(userData == 0);

      auto& state = *as_ptr<State>(userData);
      state.endpointVolume.reset();
      state.audioDevice.reset();

      if (auto result = state.deviceEnumerator->GetDefaultAudioEndpoint(
              eRender, eConsole, std::out_ptr(state.audioDevice));
          result == E_NOTFOUND)
      {
        break;
      } else {
        throwIfCOM(result);
      }

      throwIfCOM(state.audioDevice->Activate(  //
          __uuidof(IAudioEndpointVolume),
          CLSCTX_INPROC_SERVER,
          nullptr,
          std::out_ptr(state.endpointVolume)));
      throwIfCOM(state.endpointVolume->RegisterControlChangeNotify(
          ComCallback<EndpointHandler>(*state.guid, hwnd)));
      ChangeAudio(state);
      break;
    }
    case UserMessage::ChangeAudio: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      throwIf(userData == 0);

      ChangeAudio(*as_ptr<State>(userData));
      break;
    }
    case WM_COMMAND:
      switch (LOWORD(wParam)) {
        case UserMessage::TrayLicense: {
          auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
          throwIf(userData == 0);

          auto& state = *as_ptr<State>(userData);
          if (state.dialog != nullptr) {
            break;
          }

          state.dialog = CreateDialogIndirectParamW(  //
              state.hInstance,
              LicenseDialogData::get(),
              nullptr,
              DialogProc,
              reinterpret_cast<LPARAM>(&state.dialog));
          throwIf(state.dialog == nullptr);
          break;
        }
        case UserMessage::TrayExit:
          throwIf(DestroyWindow(hwnd) == 0);
          break;
      }
      break;
    case WM_CLOSE:
      throwIf(DestroyWindow(hwnd) == 0);
      break;
    case WM_DESTROY: {
      auto userData = GetWindowLongPtrW(hwnd, GWLP_USERDATA);
      throwIf(userData == 0);

      auto& state = *as_ptr<State>(userData);
      if (state.dialog != nullptr) {
        throwIf(DestroyWindow(state.dialog) == 0);
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

  (void)SetThreadDpiAwarenessContext(
      DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

  throwIfCOM(CoInitialize(nullptr));

  auto guid = GUID {};
  throwIfCOM(CoCreateGuid(&guid));

  auto deviceEnumerator = ComPtr<IMMDeviceEnumerator>();
  throwIfCOM(CoCreateInstance(  //
      __uuidof(MMDeviceEnumerator),
      nullptr,
      CLSCTX_INPROC_SERVER,
      __uuidof(IMMDeviceEnumerator),
      std::out_ptr(deviceEnumerator)));

  auto richEdit = Library(L"Riched20.dll");
  auto cursor = LoadCursorW(nullptr, IDC_ARROW);
  throwIf(cursor == nullptr);

  auto mainWindowClass = WNDCLASSW  //
      {.lpfnWndProc = MainWndProc,
       .hInstance = hInstance,
       .hCursor = cursor,
       .lpszClassName = L"AlwaysMute - Main"};
  auto mainAtom = RegisterClassW(&mainWindowClass);
  throwIf(mainAtom == 0);

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
  throwIf(window == nullptr);

  throwIfCOM(deviceEnumerator->RegisterEndpointNotificationCallback(
      ComCallback<NotificationClient>(window)));

  auto sndVol = Library(L"SndVolSSO.dll");
  auto icon = LoadIconW(sndVol.library, MAKEINTRESOURCEW(120));
  throwIf(icon == nullptr);

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
  throwIf(PostMessageW(window, UserMessage::GetDefaultEndpoint, 0, 0) == 0);

  while (true) {
    if (auto result = GetMessageW(&msg, nullptr, 0, 0); result == 0) {
      break;
    } else {
      throwIf(result == -1);
    }

    if (state.dialog != nullptr && IsDialogMessageW(state.dialog, &msg) != 0) {
      continue;
    }

    (void)DispatchMessageW(&msg);
  }

  return static_cast<int>(msg.wParam);
}

}  // namespace

int WINAPI wWinMain(  //
    _In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE,
    _In_ LPWSTR,
    _In_ int)
{
  try {
    return TryMain(hInstance);
  } catch (com_error const& error) {
    outputSystemError(static_cast<DWORD>(error.code()));
  } catch (std::exception const& error) {
    OutputDebugStringA(error.what());
    OutputDebugStringW(L"\n");
  }

  return 1;
}
