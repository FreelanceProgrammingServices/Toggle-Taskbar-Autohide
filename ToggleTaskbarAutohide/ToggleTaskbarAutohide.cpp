/*
╔═══════════════════════════════════════════════════════════════════════════════╗
║ ToggleTaskbarAutohide.cpp                                                         ║
╠═══════════════════════════════════════════════════════════════════════════════╣
║ Purpose: Toggle Windows taskbar auto-hide option with a convenient shortcut.  ║
║                                                                               ║
║ Author : Nikolas Wieczorek                                                    ║
║ Version: 1.0.0                                                                ║
║ Created: March 2025                                                           ║
╚═══════════════════════════════════════════════════════════════════════════════╝
*/

/*
=======================[ SYSTEM OVERVIEW ]=========================
Main Components:
1. Registry Manipulation: Modifies StuckRects3 registry values to control
   taskbar behavior
2. Explorer Process Handling: Kills and restarts explorer.exe to apply changes
3. Window State Preservation: Tracks and restores open Explorer windows
4. Foreground App Preservation: Remembers and restores focused application
5. System Tray Integration: Optional tray mode for persistent access

Key Functions:
- ExecuteToggleAction(): Main orchestration function
- ToggleTaskbarSetting(): Registry manipulation
- GetOpenExplorerWindows() / RestoreExplorerWindows(): Window state handling
- GetForegroundAppInfo() / RestoreForegroundApp(): Focus preservation
*/

#include "framework.h"
#include "ToggleTaskbarAutohide.h"
#include <Windows.h>
#include <winreg.h>
#include <vector>
#include <string>
#include <ShlObj.h>
#include <TlHelp32.h>
#include <atlbase.h>
#include <shellapi.h>
#include <map>
#include <Psapi.h>
#include <strsafe.h>

#define TRAY_MODE true
#define WM_TRAYICON (WM_USER + 1)
#define ID_TRAY_EXIT 1001
#define TASKBAR_ALWAYS_VISIBLE 0x02
#define TASKBAR_AUTOHIDE 0x03

/*
 * These structs store information about Explorer windows and foreground
 * applications to preserve state when restarting Explorer
 */
struct ExplorerWindow {
    std::wstring path;
    RECT position;
    WINDOWPLACEMENT placement;
    HWND hwnd;
    HWND focusedHwnd;
    DWORD zOrder;
};

struct ForegroundAppInfo {
    HWND hwnd;
    DWORD processId;
    std::wstring executablePath;
    std::wstring windowTitle;
    WINDOWPLACEMENT placement;
};

HWND g_hwnd = NULL;
NOTIFYICONDATA g_nid = { 0 };
bool g_trayMode = TRAY_MODE;
UINT WM_TASKBARCREATED = 0;
bool g_isRestartingExplorer = false;
HANDLE g_watchdogThread = NULL;

std::vector<ExplorerWindow> GetOpenExplorerWindows();
void RestoreExplorerWindows(const std::vector<ExplorerWindow>& windows);
BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam);
void KillExplorerProcess();
void StartExplorerProcess();
bool ToggleTaskbarSetting();
bool HasCommandLineOption(const wchar_t* option);
ForegroundAppInfo GetForegroundAppInfo();
void RestoreForegroundApp(const ForegroundAppInfo& appInfo);
LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
void ExecuteToggleAction();
void SetupTrayIcon(HWND hwnd);
void UpdateTrayIconTooltip();
void RemoveTrayIcon();
BYTE GetCurrentTaskbarSetting();
DWORD WINAPI WatchdogThreadProc(LPVOID lpParam);

/*
 * Watchdog Thread:
 * This thread monitors system state after Explorer restarts
 * and ensures tray icon is properly recreated
 */
DWORD WINAPI WatchdogThreadProc(LPVOID lpParam) {
    Sleep(2000);

    if (g_hwnd != NULL && IsWindow(g_hwnd)) {
        PostMessage(g_hwnd, WM_TASKBARCREATED, 0, 0);
    }
    g_isRestartingExplorer = false;
    g_watchdogThread = NULL;
    return 0;
}

/*
 * Main Action Orchestrator:
 * This function coordinates the entire toggle operation:
 * 1. Captures current foreground window and Explorer windows
 * 2. Toggles registry settings
 * 3. Restarts Explorer process
 * 4. Restores Explorer windows and focused application
 */
void ExecuteToggleAction() {
    ForegroundAppInfo foregroundApp = GetForegroundAppInfo();
    bool shouldReopenExplorer = !HasCommandLineOption(L"--noreopenexplorer");
    std::vector<ExplorerWindow> explorerWindows;
    if (shouldReopenExplorer) explorerWindows = GetOpenExplorerWindows();
    ToggleTaskbarSetting();
    g_isRestartingExplorer = true;
    if (g_trayMode && g_watchdogThread == NULL) {
        g_watchdogThread = CreateThread(NULL, 0, WatchdogThreadProc, NULL, 0, NULL);
    }
    KillExplorerProcess();
    StartExplorerProcess();
    Sleep(750);
    if (shouldReopenExplorer) RestoreExplorerWindows(explorerWindows);
    Sleep(500);
    RestoreForegroundApp(foregroundApp);
    if (g_hwnd && g_trayMode) {
        UpdateTrayIconTooltip();
        if (g_isRestartingExplorer) {
            SetTimer(g_hwnd, 1234, 2000, NULL);
        }
    }
}

/*
 * Application Entry Point:
 */
int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR lpCmdLine, _In_ int nCmdShow) {
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(nCmdShow);

    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) return 1;
    g_trayMode = HasCommandLineOption(L"--tray") || TRAY_MODE;

    if (!g_trayMode) {
        ExecuteToggleAction();
        CoUninitialize();
        return 0;
    }
    WM_TASKBARCREATED = RegisterWindowMessageW(L"TaskbarCreated");
    WNDCLASSEXW wcex = { sizeof(WNDCLASSEXW) };
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.lpszClassName = L"ToggleTaskbarAutohideClass";
    wcex.hIcon = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_APPLICATION));
    wcex.hCursor = LoadCursorW(nullptr, IDC_ARROW);

    RegisterClassExW(&wcex);
    g_hwnd = CreateWindowW(L"ToggleTaskbarAutohideClass", L"Taskbar Autohide Toggle",
        WS_OVERLAPPED, CW_USEDEFAULT, 0, 0, 0,
        HWND_MESSAGE,
        nullptr, hInstance, nullptr);

    if (!g_hwnd) {
        CoUninitialize();
        return 1;
    }

    SetupTrayIcon(g_hwnd);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    RemoveTrayIcon();
    CoUninitialize();
    return (int)msg.wParam;
}

/*
 * Foreground window state capture:
 * Captures detailed information about the currently active window
 * to restore focus after Explorer restarts
 */
ForegroundAppInfo GetForegroundAppInfo() {
    ForegroundAppInfo info = {};
    info.hwnd = GetForegroundWindow();

    if (info.hwnd) {
        GetWindowThreadProcessId(info.hwnd, &info.processId);

        if (info.processId) {
            HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, info.processId);
            if (hProcess) {
                wchar_t executablePath[MAX_PATH] = { 0 };
                if (GetModuleFileNameEx(hProcess, NULL, executablePath, MAX_PATH)) {
                    info.executablePath = executablePath;
                }
                CloseHandle(hProcess);
            }
        }

        wchar_t windowTitle[1024] = { 0 };
        GetWindowText(info.hwnd, windowTitle, 1024);
        info.windowTitle = windowTitle;

        info.placement.length = sizeof(WINDOWPLACEMENT);
        GetWindowPlacement(info.hwnd, &info.placement);
    }

    return info;
}

void RestoreForegroundApp(const ForegroundAppInfo& appInfo) {
    if (!appInfo.hwnd || !IsWindow(appInfo.hwnd)) {
        HWND foundWindow = NULL;

        struct FindWindowParams {
            HWND* resultWindow;
            const ForegroundAppInfo* appInfo;
        };

        FindWindowParams params = { &foundWindow, &appInfo };

        EnumWindows([](HWND hwnd, LPARAM lParam) -> BOOL {
            auto params = reinterpret_cast<FindWindowParams*>(lParam);

            if (!IsWindowVisible(hwnd)) return TRUE;

            DWORD processId = 0;
            GetWindowThreadProcessId(hwnd, &processId);

            if (processId == params->appInfo->processId) {
                wchar_t title[1024] = { 0 };
                GetWindowText(hwnd, title, 1024);

                if ((wcscmp(title, params->appInfo->windowTitle.c_str()) == 0) ||
                    (wcsstr(title, params->appInfo->windowTitle.c_str()) != NULL &&
                        params->appInfo->windowTitle.length() > 0)) {
                    *(params->resultWindow) = hwnd;
                    return FALSE;
                }
            }
            return TRUE;
            }, reinterpret_cast<LPARAM>(&params));

        if (foundWindow) {
            if (appInfo.placement.showCmd == SW_SHOWMAXIMIZED) {
                ShowWindow(foundWindow, SW_SHOWMAXIMIZED);
            }
            else if (appInfo.placement.showCmd == SW_SHOWMINIMIZED) {
                ShowWindow(foundWindow, SW_RESTORE);
            }
            else {
                ShowWindow(foundWindow, SW_NORMAL);
            }

            SetForegroundWindow(foundWindow);
            SetActiveWindow(foundWindow);
            SetFocus(foundWindow);
        }
    }
    else {
        if (appInfo.placement.showCmd == SW_SHOWMAXIMIZED) {
            ShowWindow(appInfo.hwnd, SW_SHOWMAXIMIZED);
        }
        else if (appInfo.placement.showCmd == SW_SHOWMINIMIZED) {
            ShowWindow(appInfo.hwnd, SW_RESTORE);
        }
        else {
            ShowWindow(appInfo.hwnd, SW_NORMAL);
        }

        SetForegroundWindow(appInfo.hwnd);
        SetActiveWindow(appInfo.hwnd);
        SetFocus(appInfo.hwnd);
    }
}

bool HasCommandLineOption(const wchar_t* option) {
    int argc;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    bool found = false;
    for (int i = 1; i < argc; i++) {
        if (_wcsicmp(argv[i], option) == 0) {
            found = true;
            break;
        }
    }
    LocalFree(argv);
    return found;
}

BYTE GetCurrentTaskbarSetting() {
    /*
    ╔═══════════════════════════════════════════════════════════════════════════════╗
    ║ StuckRects3 "Settings" Binary Structure Map (64-byte serialized configuration)║
    ╠═══════════════════════════════════════════════════════════════════════════════╣
    ║ Offset │ Size │ Purpose                                                       ║
    ╠════════╪══════╪═══════════════════════════════════════════════════════════════╣
    ║ 0x00   │ 4    │ Structure version identifier (typically 0x30,0x00,0x00,0x00)  ║
    ║ 0x04   │ 4    │ Configuration bitflags:                                       ║
    ║        │      │ • Bit 0: Taskbar position (0=bottom, 1=top, 2=left, 3=right)  ║
    ║        │      │ • Bits 1-31: Reserved for internal Windows use                ║
    ║ 0x08   │ 1    │ ► Visibility control flag:                                    ║
    ║        │      │ • 0x02 = Always visible (standard configuration)              ║
    ║        │      │ • 0x03 = Auto-hide enabled                                    ║
    ║ 0x09   │ 3    │ Reserved for future use                                       ║
    ║ 0x0C   │ 16   │ Taskbar position/dimension information                        ║
    ║ 0x1C   │ 36   │ Additional configuration data                                 ║
    ╚════════╧══════╧═══════════════════════════════════════════════════════════════╝
    */
    const wchar_t* keyPath = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects3";
    HKEY hKey;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, keyPath, 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) {
        const wchar_t* altKeyPath = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects2";
        result = RegOpenKeyExW(HKEY_CURRENT_USER, altKeyPath, 0, KEY_READ, &hKey);
        if (result != ERROR_SUCCESS) return TASKBAR_ALWAYS_VISIBLE; // Default
    }

    BYTE settings[64];
    DWORD settingsSize = sizeof(settings);
    DWORD type = REG_BINARY;
    result = RegQueryValueExW(hKey, L"Settings", NULL, &type, settings, &settingsSize);
    RegCloseKey(hKey);

    if (result != ERROR_SUCCESS) {
        return TASKBAR_ALWAYS_VISIBLE; 
    }

    return settings[0x08];
}

bool ToggleTaskbarSetting() {
    const wchar_t* keyPath = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects3";
    HKEY hKey;
    LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, keyPath, 0, KEY_READ | KEY_WRITE, &hKey);
    if (result != ERROR_SUCCESS) {
        const wchar_t* altKeyPath = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\StuckRects2";
        result = RegOpenKeyExW(HKEY_CURRENT_USER, altKeyPath, 0, KEY_READ | KEY_WRITE, &hKey);
        if (result != ERROR_SUCCESS) return false;
    }

    BYTE settings[64];
    DWORD settingsSize = sizeof(settings);
    DWORD type = REG_BINARY;
    result = RegQueryValueExW(hKey, L"Settings", NULL, &type, settings, &settingsSize);
    if (result != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return false;
    }

    settings[0x08] = (settings[0x08] == TASKBAR_ALWAYS_VISIBLE) ? TASKBAR_AUTOHIDE : TASKBAR_ALWAYS_VISIBLE;
    result = RegSetValueExW(hKey, L"Settings", 0, REG_BINARY, settings, settingsSize);
    RegCloseKey(hKey);

    SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 0,
        (LPARAM)L"TraySettings", SMTO_ABORTIFHUNG, 1000, NULL);

    return (result == ERROR_SUCCESS);
}

void KillExplorerProcess() {
    EnumWindows([](HWND hwnd, LPARAM lParam) -> BOOL {
        wchar_t className[256];
        GetClassName(hwnd, className, 256);
        if (wcscmp(className, L"CabinetWClass") == 0) PostMessage(hwnd, WM_CLOSE, 0, 0);
        return TRUE;
        }, 0);

    Sleep(150);

    DWORD ourProcessId = GetCurrentProcessId();
    HANDLE hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0);
    PROCESSENTRY32 pEntry;
    pEntry.dwSize = sizeof(pEntry);
    BOOL hRes = Process32First(hSnapShot, &pEntry);
    while (hRes) {
        if (wcscmp(pEntry.szExeFile, L"explorer.exe") == 0 && pEntry.th32ProcessID != ourProcessId) {
            HANDLE hProcess = OpenProcess(PROCESS_TERMINATE, 0, pEntry.th32ProcessID);
            if (hProcess != NULL) {
                TerminateProcess(hProcess, 0);
                CloseHandle(hProcess);
            }
        }
        hRes = Process32Next(hSnapShot, &pEntry);
    }
    CloseHandle(hSnapShot);
}

void StartExplorerProcess() {
    STARTUPINFOW si = { sizeof(STARTUPINFOW) };
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));
    if (CreateProcessW(L"C:\\Windows\\explorer.exe", NULL, NULL, NULL, FALSE,
        0, NULL, NULL, &si, &pi)) {
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
    }
}

std::vector<ExplorerWindow> GetOpenExplorerWindows() {
    std::vector<ExplorerWindow> windows;
    std::map<DWORD, ExplorerWindow> windowsByZOrder;
    HWND focusedWindow = GetForegroundWindow();
    HWND hwnd = GetTopWindow(NULL);
    DWORD zOrder = 0;

    while (hwnd) {
        if (IsWindowVisible(hwnd)) {
            wchar_t className[256];
            GetClassName(hwnd, className, 256);

            if (wcscmp(className, L"CabinetWClass") == 0) {
                ExplorerWindow window;
                window.hwnd = hwnd;
                window.focusedHwnd = NULL;
                window.zOrder = zOrder;
                window.placement.length = sizeof(WINDOWPLACEMENT);
                GetWindowPlacement(hwnd, &window.placement);
                window.position = window.placement.rcNormalPosition;

                if (hwnd == focusedWindow) window.focusedHwnd = hwnd;
                else {
                    HWND childFocus = GetFocus();
                    if (childFocus && IsChild(hwnd, childFocus)) window.focusedHwnd = childFocus;
                }

                CComPtr<IShellWindows> shellWindows;
                HRESULT hr = shellWindows.CoCreateInstance(CLSID_ShellWindows);
                if (SUCCEEDED(hr)) {
                    VARIANT v;
                    V_VT(&v) = VT_I4;
                    long count;
                    shellWindows->get_Count(&count);

                    for (long i = 0; i < count; i++) {
                        CComPtr<IDispatch> disp;
                        V_I4(&v) = i;
                        hr = shellWindows->Item(v, &disp);
                        if (SUCCEEDED(hr) && disp) {
                            CComPtr<IWebBrowserApp> webApp;
                            hr = disp->QueryInterface(IID_IWebBrowserApp, (void**)&webApp);
                            if (SUCCEEDED(hr) && webApp) {
                                HWND browserHwnd;
                                webApp->get_HWND((SHANDLE_PTR*)&browserHwnd);
                                if (browserHwnd == hwnd) {
                                    CComPtr<IServiceProvider> sp;
                                    hr = webApp->QueryInterface(IID_IServiceProvider, (void**)&sp);
                                    if (SUCCEEDED(hr) && sp) {
                                        CComPtr<IShellBrowser> browser;
                                        hr = sp->QueryService(SID_STopLevelBrowser, IID_IShellBrowser, (void**)&browser);
                                        if (SUCCEEDED(hr) && browser) {
                                            CComPtr<IShellView> view;
                                            hr = browser->QueryActiveShellView(&view);
                                            if (SUCCEEDED(hr) && view) {
                                                CComPtr<IFolderView> folderView;
                                                hr = view->QueryInterface(IID_IFolderView, (void**)&folderView);
                                                if (SUCCEEDED(hr) && folderView) {
                                                    CComPtr<IPersistFolder2> folder;
                                                    hr = folderView->GetFolder(IID_IPersistFolder2, (void**)&folder);
                                                    if (SUCCEEDED(hr) && folder) {
                                                        LPITEMIDLIST pidl;
                                                        hr = folder->GetCurFolder(&pidl);
                                                        if (SUCCEEDED(hr) && pidl) {
                                                            wchar_t path[MAX_PATH];
                                                            SHGetPathFromIDList(pidl, path);
                                                            window.path = path;
                                                            CoTaskMemFree(pidl);
                                                            windowsByZOrder[zOrder] = window;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        hwnd = GetNextWindow(hwnd, GW_HWNDNEXT);
        zOrder++;
    }

    for (const auto& pair : windowsByZOrder) windows.push_back(pair.second);
    return windows;
}

BOOL CALLBACK EnumWindowsProc(HWND hwnd, LPARAM lParam) {
    std::vector<ExplorerWindow>* windows = reinterpret_cast<std::vector<ExplorerWindow>*>(lParam);
    if (!IsWindowVisible(hwnd)) return TRUE;

    wchar_t className[256];
    GetClassName(hwnd, className, 256);
    if (wcscmp(className, L"CabinetWClass") != 0) return TRUE;

    ExplorerWindow window;
    window.hwnd = hwnd;
    window.placement.length = sizeof(WINDOWPLACEMENT);
    GetWindowPlacement(hwnd, &window.placement);
    window.position = window.placement.rcNormalPosition;
    HWND focusedWindow = GetForegroundWindow();

    if (hwnd == focusedWindow) window.focusedHwnd = hwnd;
    else {
        HWND childFocus = GetFocus();
        if (childFocus && IsChild(hwnd, childFocus)) window.focusedHwnd = childFocus;
        else window.focusedHwnd = NULL;
    }

    CComPtr<IShellWindows> shellWindows;
    HRESULT hr = shellWindows.CoCreateInstance(CLSID_ShellWindows);
    if (SUCCEEDED(hr)) {
        VARIANT v;
        V_VT(&v) = VT_I4;
        long count;
        shellWindows->get_Count(&count);

        for (long i = 0; i < count; i++) {
            CComPtr<IDispatch> disp;
            V_I4(&v) = i;
            hr = shellWindows->Item(v, &disp);
            if (SUCCEEDED(hr) && disp) {
                CComPtr<IWebBrowserApp> webApp;
                hr = disp->QueryInterface(IID_IWebBrowserApp, (void**)&webApp);
                if (SUCCEEDED(hr) && webApp) {
                    HWND browserHwnd;
                    webApp->get_HWND((SHANDLE_PTR*)&browserHwnd);
                    if (browserHwnd == hwnd) {
                        CComPtr<IServiceProvider> sp;
                        hr = webApp->QueryInterface(IID_IServiceProvider, (void**)&sp);
                        if (SUCCEEDED(hr) && sp) {
                            CComPtr<IShellBrowser> browser;
                            hr = sp->QueryService(SID_STopLevelBrowser, IID_IShellBrowser, (void**)&browser);
                            if (SUCCEEDED(hr) && browser) {
                                CComPtr<IShellView> view;
                                hr = browser->QueryActiveShellView(&view);
                                if (SUCCEEDED(hr) && view) {
                                    CComPtr<IFolderView> folderView;
                                    hr = view->QueryInterface(IID_IFolderView, (void**)&folderView);
                                    if (SUCCEEDED(hr) && folderView) {
                                        CComPtr<IPersistFolder2> folder;
                                        hr = folderView->GetFolder(IID_IPersistFolder2, (void**)&folder);
                                        if (SUCCEEDED(hr) && folder) {
                                            LPITEMIDLIST pidl;
                                            hr = folder->GetCurFolder(&pidl);
                                            if (SUCCEEDED(hr) && pidl) {
                                                wchar_t path[MAX_PATH];
                                                SHGetPathFromIDList(pidl, path);
                                                window.path = path;
                                                CoTaskMemFree(pidl);
                                                windows->push_back(window);
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    return TRUE;
}

void RestoreExplorerWindows(const std::vector<ExplorerWindow>& windows) {
    std::map<std::wstring, HWND> openedPaths;
    for (auto it = windows.rbegin(); it != windows.rend(); ++it) {
        const auto& window = *it;
        if (!window.path.empty()) {
            ShellExecuteW(NULL, L"open", L"explorer.exe", window.path.c_str(), NULL, SW_SHOWNORMAL);
            Sleep(500);
            HWND newHwnd = NULL;
            struct FindParams {
                HWND* result;
                const std::wstring* targetPath;
            };
            FindParams findParams = { &newHwnd, &window.path };
            EnumWindows([](HWND hwnd, LPARAM lParam) -> BOOL {
                FindParams* params = reinterpret_cast<FindParams*>(lParam);
                wchar_t className[256];
                GetClassName(hwnd, className, 256);
                if (wcscmp(className, L"CabinetWClass") == 0 && IsWindowVisible(hwnd)) {
                    std::wstring path;
                    CComPtr<IShellWindows> shellWindows;
                    HRESULT hr = shellWindows.CoCreateInstance(CLSID_ShellWindows);
                    if (SUCCEEDED(hr)) {
                        VARIANT v;
                        V_VT(&v) = VT_I4;
                        long count;
                        shellWindows->get_Count(&count);
                        for (long i = 0; i < count; i++) {
                            CComPtr<IDispatch> disp;
                            V_I4(&v) = i;
                            hr = shellWindows->Item(v, &disp);
                            if (SUCCEEDED(hr) && disp) {
                                CComPtr<IWebBrowserApp> webApp;
                                hr = disp->QueryInterface(IID_IWebBrowserApp, (void**)&webApp);
                                if (SUCCEEDED(hr) && webApp) {
                                    HWND browserHwnd;
                                    webApp->get_HWND((SHANDLE_PTR*)&browserHwnd);
                                    if (browserHwnd == hwnd) {
                                        CComPtr<IServiceProvider> sp;
                                        hr = webApp->QueryInterface(IID_IServiceProvider, (void**)&sp);
                                        if (SUCCEEDED(hr) && sp) {
                                            CComPtr<IShellBrowser> browser;
                                            hr = sp->QueryService(SID_STopLevelBrowser, IID_IShellBrowser, (void**)&browser);
                                            if (SUCCEEDED(hr) && browser) {
                                                CComPtr<IShellView> view;
                                                hr = browser->QueryActiveShellView(&view);
                                                if (SUCCEEDED(hr) && view) {
                                                    CComPtr<IFolderView> folderView;
                                                    hr = view->QueryInterface(IID_IFolderView, (void**)&folderView);
                                                    if (SUCCEEDED(hr) && folderView) {
                                                        CComPtr<IPersistFolder2> folder;
                                                        hr = folderView->GetFolder(IID_IPersistFolder2, (void**)&folder);
                                                        if (SUCCEEDED(hr) && folder) {
                                                            LPITEMIDLIST pidl;
                                                            hr = folder->GetCurFolder(&pidl);
                                                            if (SUCCEEDED(hr) && pidl) {
                                                                wchar_t pathBuffer[MAX_PATH];
                                                                SHGetPathFromIDList(pidl, pathBuffer);
                                                                path = pathBuffer;
                                                                CoTaskMemFree(pidl);
                                                                if (_wcsicmp(path.c_str(), params->targetPath->c_str()) == 0) {
                                                                    *(params->result) = hwnd;
                                                                    return FALSE;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                return TRUE;
                }, (LPARAM)&findParams);

            if (newHwnd) {
                openedPaths[window.path] = newHwnd;
                ShowWindow(newHwnd, SW_NORMAL);
                Sleep(200);
                MoveWindow(
                    newHwnd,
                    window.position.left,
                    window.position.top,
                    window.position.right - window.position.left,
                    window.position.bottom - window.position.top,
                    TRUE
                );
                if (window.placement.showCmd == SW_MAXIMIZE) {
                    ShowWindow(newHwnd, SW_MAXIMIZE);
                }
                else if (window.placement.showCmd == SW_MINIMIZE) {
                    ShowWindow(newHwnd, SW_MINIMIZE);
                }
                HWND insertAfter = HWND_TOP;
                if (window.placement.showCmd == SW_MINIMIZE) {
                    insertAfter = HWND_BOTTOM;
                }
                SetWindowPos(newHwnd, insertAfter, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            }
        }
    }
}

/*
 * Tray Icon Tooltip Update:
 * Updates system tray tooltip to reflect current taskbar state
 */
void UpdateTrayIconTooltip() {
    BYTE currentSetting = GetCurrentTaskbarSetting();
    g_nid.uFlags = NIF_TIP;

    if (currentSetting == TASKBAR_AUTOHIDE) {
        wcscpy_s(g_nid.szTip, L"Show the taskbar automatically");
    }
    else {
        wcscpy_s(g_nid.szTip, L"Hide the taskbar automatically");
    }

    Shell_NotifyIcon(NIM_MODIFY, &g_nid);
}

/*
 * Tray Icon Setup:
 * Creates and configures persistent tray icon
 */
void SetupTrayIcon(HWND hwnd) {
    ZeroMemory(&g_nid, sizeof(g_nid));

    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = hwnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = (HICON)LoadImageW(GetModuleHandleW(NULL), L"tray16.ico", IMAGE_ICON, 16, 16, LR_LOADFROMFILE);
    if (!g_nid.hIcon) {
        g_nid.hIcon = LoadIconW(GetModuleHandleW(NULL), MAKEINTRESOURCEW(IDI_APPLICATION));
    }
    BYTE currentSetting = GetCurrentTaskbarSetting();
    if (currentSetting == TASKBAR_AUTOHIDE) {
        wcscpy_s(g_nid.szTip, L"Show the taskbar automatically");
    }
    else {
        wcscpy_s(g_nid.szTip, L"Hide the taskbar automatically");
    }
    Shell_NotifyIconW(NIM_ADD, &g_nid);
}

void RemoveTrayIcon() {
    if (g_nid.cbSize > 0) {
        Shell_NotifyIcon(NIM_DELETE, &g_nid);
        if (g_nid.hIcon) {
            DestroyIcon(g_nid.hIcon);
            g_nid.hIcon = NULL;
        }
    }
}

/*
 * Window Procedure:
 * Handles window and system tray messages for tray mode operation
 */
LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {
    if (message == WM_TASKBARCREATED && g_trayMode) {
        SetupTrayIcon(hwnd);
        return 0;
    }

    switch (message) {
    case WM_TIMER:
        if (wParam == 1234) {
            KillTimer(hwnd, 1234);
            if (g_trayMode) {
                RemoveTrayIcon();
                SetupTrayIcon(hwnd);
                g_isRestartingExplorer = false;
            }
            return 0;
        }
        break;

    case WM_TRAYICON:
        if (lParam == WM_LBUTTONUP) {
            ExecuteToggleAction();
            return 0;
        }
        else if (lParam == WM_RBUTTONUP) {
            POINT pt;
            GetCursorPos(&pt);
            HMENU hMenu = CreatePopupMenu();
            if (hMenu) {
                InsertMenu(hMenu, -1, MF_BYPOSITION | MF_STRING, ID_TRAY_EXIT, L"&Quit application");
                SetForegroundWindow(hwnd);
                TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN,
                    pt.x, pt.y, 0, hwnd, NULL);
                PostMessage(hwnd, WM_NULL, 0, 0); 
                DestroyMenu(hMenu);
            }
            return 0;
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == ID_TRAY_EXIT) {
            PostQuitMessage(0);
            return 0;
        }
        break;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(hwnd, message, wParam, lParam);
}