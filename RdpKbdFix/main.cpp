#include <windows.h>
#include <TlHelp32.h>
#include <atlbase.h>

#include <iostream>
#include <format>
#include <unordered_map>

#define VERSION_STR                 L"1.2-ru3"

#ifdef _WIN64
#define MUTEX_NAME                  L"Global\\LowLevelKeyboardHookFix64"
#else
#define MUTEX_NAME                  L"Global\\LowLevelKeyboardHookFix32"
#endif

#define POLL_INTERVAL_MS            (2500)

static const wchar_t* g_pwszProcsToMonitor[] = { 
#ifndef _WIN64
    L"vmware-remotemks.exe",
#else
    L"vmware-vmx.exe"
#endif
};

static HMODULE g_hSelf = nullptr;
static HOOKPROC g_pVMwareHookFunc = nullptr;
static HKL g_hklCurrentKeyboard = nullptr;
static HKL g_hklRussian = nullptr;        // Russian keyboard layout handle (if installed)
static HKL g_hklEnglish = nullptr;        // English keyboard layout handle (if installed)
static bool g_bVMInRussianMode = false;   // Tracks assumed layout state inside the VM
static bool g_bReportedShiftDown = false; // Whether we've told VMware that Shift is held
static bool g_bUserAltDown = false;       // Tracks real (non-injected) Alt key state
static bool g_bUserShiftDown = false;     // Tracks real (non-injected) Shift key state

static void ExitPrompt(const std::wstring& wstrMsg=L"")
{
    std::wcout << std::format(L"{}\nPress ENTER to exit\n", wstrMsg);
    std::cin.get();
}

static bool SetPrivilege(HANDLE hToken,
                         LPCWSTR lpszPrivilege,
                         bool bEnablePrivilege)
{
    TOKEN_PRIVILEGES tp;
    LUID luid;

    if (!LookupPrivilegeValueW(nullptr, lpszPrivilege, &luid))
    {
        std::wcout << std::format(L"LookupPrivilegeValueW failed with {}\n", GetLastError());
        return false;
    }

    tp.PrivilegeCount = 1;
    tp.Privileges[0].Luid = luid;
    tp.Privileges[0].Attributes = bEnablePrivilege ? SE_PRIVILEGE_ENABLED : 0;

    if (!AdjustTokenPrivileges(hToken,
                               FALSE,
                               &tp,
                               sizeof(TOKEN_PRIVILEGES),
                               nullptr,
                               nullptr)
        || GetLastError() == ERROR_NOT_ALL_ASSIGNED)
    {
        std::wcout << std::format(L"AdjustTokenPrivileges failed with {}\n", GetLastError());
        return false;
    }

    return true;
}

// Returns true if the character is Cyrillic (requires Russian keyboard layout)
static bool IsCyrillicChar(WCHAR wc)
{
    return (wc >= 0x0400 && wc <= 0x04FF);
}

// Returns true if the character is a Latin letter (requires English keyboard layout)
static bool IsLatinChar(WCHAR wc)
{
    return (wc >= L'A' && wc <= L'Z') || (wc >= L'a' && wc <= L'z');
}

// Fallback: search all installed keyboard layouts for the given Unicode char.
static SHORT VkKeyScanExWAllLayouts(WCHAR wc)
{
    SHORT result = VkKeyScanExW(wc, g_hklCurrentKeyboard);
    if (result != -1 && (result & 0xFF) != 0xFF)
        return result;

    HKL layouts[64] = {};
    int nLayouts = GetKeyboardLayoutList(_countof(layouts), layouts);
    for (int i = 0; i < nLayouts; i++)
    {
        if (layouts[i] == g_hklCurrentKeyboard) continue;
        result = VkKeyScanExW(wc, layouts[i]);
        if (result != -1 && (result & 0xFF) != 0xFF)
            return result;
    }
    return -1;
}

static void TranslateVKPacket(WPARAM wParam, PKBDLLHOOKSTRUCT pkbdStruct)
{
    INPUT inputs[8] = {};
    int inputCount = 0;

    WCHAR wc = static_cast<WCHAR>(pkbdStruct->scanCode);

    // Auto-switch VM keyboard layout when the character language changes.
    // Sends Left Alt + Left Shift to toggle layout inside the VM, so the user
    // can type both Russian and English from the phone regardless of which
    // layout the VM currently has active.
    bool bNeedRussian = IsCyrillicChar(wc);
    bool bNeedEnglish = IsLatinChar(wc);
    bool bNeedSwitch  = (bNeedRussian && !g_bVMInRussianMode) ||
                        (bNeedEnglish &&  g_bVMInRussianMode);

    if (bNeedSwitch)
    {
        // Left Alt down
        inputs[inputCount].type = INPUT_KEYBOARD;
        inputs[inputCount].ki.wScan = static_cast<WORD>(MapVirtualKeyA(VK_LMENU, MAPVK_VK_TO_VSC));
        inputs[inputCount++].ki.dwFlags = KEYEVENTF_SCANCODE;

        // Left Shift down  (Alt+Shift = Windows layout switch shortcut)
        inputs[inputCount].type = INPUT_KEYBOARD;
        inputs[inputCount].ki.wScan = static_cast<WORD>(MapVirtualKeyA(VK_LSHIFT, MAPVK_VK_TO_VSC));
        inputs[inputCount++].ki.dwFlags = KEYEVENTF_SCANCODE;

        // Left Shift up
        inputs[inputCount].type = INPUT_KEYBOARD;
        inputs[inputCount].ki.wScan = static_cast<WORD>(MapVirtualKeyA(VK_LSHIFT, MAPVK_VK_TO_VSC));
        inputs[inputCount++].ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;

        // Left Alt up
        inputs[inputCount].type = INPUT_KEYBOARD;
        inputs[inputCount].ki.wScan = static_cast<WORD>(MapVirtualKeyA(VK_LMENU, MAPVK_VK_TO_VSC));
        inputs[inputCount++].ki.dwFlags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP;

        g_bVMInRussianMode = bNeedRussian;
        // Alt+Shift releases any Shift we had previously pressed for uppercase letters
        g_bReportedShiftDown = false;
    }

    // Pick the best HKL for this character (prefer the language-specific one)
    HKL hklPreferred = bNeedRussian ? (g_hklRussian ? g_hklRussian : g_hklCurrentKeyboard)
                                    : (g_hklEnglish ? g_hklEnglish : g_hklCurrentKeyboard);

    SHORT vkResult = VkKeyScanExW(wc, hklPreferred);
    if (vkResult == -1 || (vkResult & 0xFF) == 0xFF)
        vkResult = VkKeyScanExWAllLayouts(wc);
    if (vkResult == -1)
        return; // Character not found in any installed layout

    // HIBYTE bit 0 = Shift required (works for any language, including Cyrillic uppercase)
    bool bNeedsShift = (HIBYTE(vkResult) & 1) != 0;

    if (bNeedsShift != g_bReportedShiftDown)
    {
        inputs[inputCount].type = INPUT_KEYBOARD;
        inputs[inputCount].ki.wScan = static_cast<WORD>(MapVirtualKeyA(VK_LSHIFT, MAPVK_VK_TO_VSC));
        inputs[inputCount].ki.dwFlags = KEYEVENTF_SCANCODE;
        if (g_bReportedShiftDown)
            inputs[inputCount].ki.dwFlags |= KEYEVENTF_KEYUP;
        ++inputCount;
        g_bReportedShiftDown = bNeedsShift;
    }

    // Append the character scancode
    inputs[inputCount].type = INPUT_KEYBOARD;
    inputs[inputCount].ki.wScan = static_cast<WORD>(MapVirtualKeyA(vkResult & 0xFF, MAPVK_VK_TO_VSC));
    inputs[inputCount].ki.dwFlags = KEYEVENTF_SCANCODE;
    if (wParam == WM_KEYUP)
        inputs[inputCount].ki.dwFlags |= KEYEVENTF_KEYUP;
    ++inputCount;

    // Send everything atomically in one call so the layout switch and character
    // are delivered to the VM in the correct order without interleaving.
    SendInput(inputCount, inputs, sizeof(INPUT));
}

static LRESULT __stdcall HookFunc(int nCode, WPARAM wParam, LPARAM lParam)
{
    PKBDLLHOOKSTRUCT pkbdStruct = reinterpret_cast<PKBDLLHOOKSTRUCT>(lParam);

    if (nCode >= 0 && (wParam == WM_KEYDOWN || wParam == WM_KEYUP))
    {
        bool bIsInjected = (pkbdStruct->flags & LLKHF_INJECTED) != 0;

        if (pkbdStruct->vkCode == VK_PACKET)
        {
            TranslateVKPacket(wParam, pkbdStruct);
            return 1;
        }

        // Only process real (non-injected) keystrokes below.
        // Injected events are our own SendInput calls — skip to avoid double-toggling.
        if (!bIsInjected)
        {
            // ── Scroll Lock: manual layout sync ─────────────────────────────────
            // Press Scroll Lock to toggle RdpKbdFix's idea of the VM's current layout.
            // Use this whenever the DLL gets out of sync (e.g. after connecting,
            // or after manually switching layout via the on-screen function keys).
            // The key is blocked from reaching VMware.
            // Audio feedback: one beep = now tracking English, two beeps = Russian.
            if (pkbdStruct->vkCode == VK_SCROLL && wParam == WM_KEYDOWN)
            {
                g_bVMInRussianMode  = !g_bVMInRussianMode;
                g_bReportedShiftDown = false;
                if (g_bVMInRussianMode)
                {
                    MessageBeep(MB_ICONINFORMATION); // one distinct beep = Russian
                    Sleep(180);
                    MessageBeep(MB_ICONINFORMATION); // two beeps
                }
                else
                {
                    MessageBeep(MB_OK); // one short beep = English
                }
                return 1; // Block key — do not pass to VMware
            }

            // ── Track real Alt state ─────────────────────────────────────────────
            if (pkbdStruct->vkCode == VK_LMENU || pkbdStruct->vkCode == VK_RMENU)
                g_bUserAltDown = (wParam == WM_KEYDOWN);

            // ── Track real Shift state + detect manual Alt+Shift ─────────────────
            // When the user manually presses Alt+Shift (e.g. via function row on iPhone
            // or from a physical keyboard), the VM's layout switches but our
            // g_bVMInRussianMode variable isn't updated — causing the next VK_PACKET
            // to trigger a spurious layout switch that corrupts Shift state.
            // Fix: detect real Alt+Shift and mirror the toggle in our tracking variable.
            if (pkbdStruct->vkCode == VK_LSHIFT || pkbdStruct->vkCode == VK_RSHIFT)
            {
                bool bWasDown = g_bUserShiftDown;
                g_bUserShiftDown = (wParam == WM_KEYDOWN);

                if (wParam == WM_KEYDOWN && !bWasDown && g_bUserAltDown)
                {
                    // Real Alt+Shift: VM layout just toggled, update our tracking
                    g_bVMInRussianMode   = !g_bVMInRussianMode;
                    g_bReportedShiftDown = false;
                }
            }
        }
    }

    return g_pVMwareHookFunc(nCode, wParam, lParam);
}

static LRESULT __stdcall HookFuncGlobal(int nCode, WPARAM wParam, LPARAM lParam)
{
    PKBDLLHOOKSTRUCT pkbdStruct = reinterpret_cast<PKBDLLHOOKSTRUCT>(lParam);

    if (nCode >= 0                                      &&
        (wParam == WM_KEYDOWN || wParam == WM_KEYUP)    &&
        pkbdStruct->vkCode == VK_PACKET)
    {
        TranslateVKPacket(wParam, pkbdStruct);
        return 1;
    }

    return CallNextHookEx(nullptr, nCode, wParam, lParam);
}

static HHOOK WINAPI _SetWindowsHookExA(int idHook, HOOKPROC lpfn, HINSTANCE hmod, DWORD dwThreadId)
{
    if (idHook != WH_KEYBOARD_LL)
    {
        return SetWindowsHookExA(idHook, lpfn, hmod, dwThreadId);
    }

    g_pVMwareHookFunc = lpfn;
    return SetWindowsHookExW(WH_KEYBOARD_LL, HookFunc, g_hSelf, 0);
}

static void PerformFix()
{
    UINT_PTR pTargetDll = reinterpret_cast<UINT_PTR>(GetModuleHandleW(nullptr));
    if (!pTargetDll)
    {
        OutputDebugStringW(L"Unable to get module\n");
        return;
    }

    PIMAGE_DOS_HEADER pDosHeader = reinterpret_cast<PIMAGE_DOS_HEADER>(pTargetDll);
    PIMAGE_NT_HEADERS pNtHeader = reinterpret_cast<PIMAGE_NT_HEADERS>(pTargetDll + pDosHeader->e_lfanew);
    PIMAGE_IMPORT_DESCRIPTOR pImportDescriptor = reinterpret_cast<PIMAGE_IMPORT_DESCRIPTOR>(pTargetDll + pNtHeader->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT].VirtualAddress);

    while (pImportDescriptor->Name)
    {
        LPCSTR pszCurrentLibName = reinterpret_cast<LPCSTR>(pTargetDll + pImportDescriptor->Name);

        if (_stricmp(pszCurrentLibName, "user32.dll") == 0)
        {
            PIMAGE_THUNK_DATA pOrigFirstThunk = reinterpret_cast<PIMAGE_THUNK_DATA>(pTargetDll + pImportDescriptor->OriginalFirstThunk);
            PIMAGE_THUNK_DATA pFirstThunk = reinterpret_cast<PIMAGE_THUNK_DATA>(pTargetDll + pImportDescriptor->FirstThunk);

            while (pOrigFirstThunk->u1.AddressOfData)
            {
                PIMAGE_IMPORT_BY_NAME pFuncName = reinterpret_cast<PIMAGE_IMPORT_BY_NAME>(pTargetDll + pOrigFirstThunk->u1.AddressOfData);
                
                if (strcmp(pFuncName->Name, "SetWindowsHookExA") == 0)
                {
                    DWORD dwOldProtect = 0;
                    if (!VirtualProtect(&pFirstThunk->u1.Function, sizeof(PVOID), PAGE_READWRITE, &dwOldProtect))
                    {
                        OutputDebugStringW(L"VirtualProtect failure\n");
                        return;
                    }

                    pFirstThunk->u1.Function = (UINT_PTR)_SetWindowsHookExA;

                    if (!VirtualProtect(&pFirstThunk->u1.Function, sizeof(PVOID), dwOldProtect, &dwOldProtect))
                    {
                        OutputDebugStringW(L"VirtualProtect failure\n");
                        return;
                    }
                    
                    return;
                }

                ++pOrigFirstThunk;
                ++pFirstThunk;
            }

            break;
        }

        ++pImportDescriptor;
    }

    OutputDebugStringW(L"Unable to find IAT hook target\n");
    return;
}

static bool IsAlreadyInjected(DWORD dwPid, PWCHAR pwszModFileName)
{
    try
    {
        MODULEENTRY32W mod32 = { 0 };
        CHandle hMods = CHandle(CreateToolhelp32Snapshot(TH32CS_SNAPMODULE | TH32CS_SNAPMODULE32, dwPid));

        mod32.dwSize = sizeof(mod32);

        if (!Module32FirstW(hMods, &mod32))
        {
            std::wcout << std::format(L"Module32FirstW failed with {}\n", GetLastError());
            return false;
        }

        do
        {
            if (_wcsicmp(mod32.szExePath, pwszModFileName) == 0)
            {
                return true;
            }
        } while (Module32NextW(hMods, &mod32));

    }
    catch (...)
    {
        std::wcout << L"Exception in IsAlreadyInjected() function\n";
    }

    return false;
}

static HRESULT Inject(HANDLE hTargetProc, PWCHAR pwszModFileName)
{
    HRESULT hRes = S_OK;
    size_t szPathByteCount = 0;
    PVOID pRemoteMem = nullptr;
    HANDLE hRemoteThread = nullptr;
    DWORD dwPid = GetProcessId(hTargetProc);

    szPathByteCount = (wcslen(pwszModFileName) + 1) * sizeof(WCHAR);
    
    pRemoteMem = VirtualAllocEx(hTargetProc, nullptr, szPathByteCount, MEM_COMMIT, PAGE_READWRITE);
    if (!pRemoteMem)
    {
        std::wcout << std::format(L"VirtualAllocEx failed with {}\n", GetLastError());
        return E_FAIL;
    }

    do
    {
        if (!WriteProcessMemory(hTargetProc, pRemoteMem, pwszModFileName, szPathByteCount, nullptr))
        {
            std::wcout << std::format(L"WriteProcessMemory failed with {}\n", GetLastError());
            hRes = E_FAIL;
            break;
        }

        PTHREAD_START_ROUTINE pLoadLibraryW = reinterpret_cast<PTHREAD_START_ROUTINE>(GetProcAddress(GetModuleHandleW(L"kernel32.dll"), "LoadLibraryW"));
        if (!pLoadLibraryW)
        {
            std::wcout << std::format(L"GetProcAddress failed with {}\n", GetLastError());
            hRes = E_FAIL;
            break;
        }

        hRemoteThread = CreateRemoteThread(hTargetProc, nullptr, 0, pLoadLibraryW, pRemoteMem, 0, nullptr);
        if (!hRemoteThread)
        {
            std::wcout << std::format(L"CreateRemoteThread failed with {}\n", GetLastError());
            hRes = E_FAIL;
            break;
        }
    } while (0);
    
    if (hRemoteThread)
    {
        WaitForSingleObject(hRemoteThread, INFINITE);
        CloseHandle(hRemoteThread);
    }

    if (pRemoteMem)
    {
        VirtualFreeEx(hTargetProc, pRemoteMem, 0, MEM_RELEASE);
    }

    if (SUCCEEDED(hRes))
    {
        return IsAlreadyInjected(dwPid, pwszModFileName) ? S_OK : E_FAIL;
    }
    else
    {
        return hRes;
    }
}

DWORD WINAPI GlobalHookThread(LPVOID lpParam)
{
    HHOOK hhkHook = SetWindowsHookExW(WH_KEYBOARD_LL, HookFuncGlobal, nullptr, 0);

    MSG msg = { 0 };
    while (GetMessageW(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    UnhookWindowsHookEx(hhkHook);

    return 0;
}

static void Entry(bool bWithGlobalHook)
{
    try
    {
        CHandle hGlobalThread;
        WCHAR wszModFileName[MAX_PATH] = { 0 };
        PROCESSENTRY32W proc32 = { 0 };
        std::unordered_map<DWORD, CHandle> umapProcsMonitored;
        FILE* fileTmp = nullptr;

        proc32.dwSize = sizeof(proc32);

        AllocConsole();
        freopen_s(&fileTmp, "CONIN$", "r", stdin);
        freopen_s(&fileTmp, "CONOUT$", "w", stdout);
        freopen_s(&fileTmp, "CONOUT$", "w", stderr);

        std::wcout << std::format(L"RdpKbdFix Version {}\n\n", VERSION_STR);

        CHandle hMutex = CHandle(CreateMutexW(nullptr, TRUE, MUTEX_NAME));
        if (!hMutex || GetLastError() == ERROR_ALREADY_EXISTS)
        {
            ExitPrompt(L"LowLevelkeyboardHookFix already running. Aborting.\n");
            return;
        }

        HANDLE hThisProcToken;
        if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES, &hThisProcToken))
        {
            ExitPrompt(std::format(L"OpenProcessToken failed with {}\n", GetLastError()));
            return;
        }

        if (!SetPrivilege(hThisProcToken, SE_DEBUG_NAME, true))
        {
            ExitPrompt(L"Hint: Run as Admin / Check Permissions\n");
            CloseHandle(hThisProcToken);
            return;
        }
        CloseHandle(hThisProcToken);

        if (!GetModuleFileNameW(g_hSelf, wszModFileName, _countof(wszModFileName)))
        {
            ExitPrompt(std::format(L"GetModuleFileNameW failed with {}\n", GetLastError()));
            return;
        }

        if (bWithGlobalHook)
        {
            hGlobalThread.Attach(CreateThread(nullptr, 0, GlobalHookThread, nullptr, 0, nullptr));
            std::wcout << L"Global hook mode ENABLED\n\n";
        }

        std::wcout << L"Monitoring to fix the following "
#ifdef _WIN64
            L"64-bit"
#else
            L"32-bit"
#endif
            " processes:\n";
        for (const wchar_t* pwszCurrentProc : g_pwszProcsToMonitor)
        {
            std::wcout << std::format(L"- `{}`\n", pwszCurrentProc);
        }

        for (;; Sleep(POLL_INTERVAL_MS))
        {
            try
            {
                CHandle hProcs = CHandle(CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0));
            
                if (!Process32FirstW(hProcs, &proc32))
                {
                    std::wcout << std::format(L"Process32FirstW failed with {}\n", GetLastError());
                    continue;
                }

                do
                {
                    for (const wchar_t* pwszCurrentProc : g_pwszProcsToMonitor)
                    {
                        if (_wcsicmp(proc32.szExeFile, pwszCurrentProc) == 0)
                        {
                            if (umapProcsMonitored.find(proc32.th32ProcessID) == umapProcsMonitored.end())
                            {
                                HANDLE hProc = OpenProcess(PROCESS_QUERY_INFORMATION |
                                                           PROCESS_CREATE_THREAD     |
                                                           PROCESS_VM_OPERATION      |
                                                           PROCESS_VM_WRITE,
                                                           FALSE, proc32.th32ProcessID);
                                if (!hProc)
                                {
                                    std::wcout << std::format(L"OpenProcess failed on {} / {} with {}\n", pwszCurrentProc, proc32.th32ProcessID, GetLastError());
                                    continue;
                                }
                            
                                if (!IsAlreadyInjected(proc32.th32ProcessID, wszModFileName))
                                {
                                    bool bInject = false;
                                    BOOL bIsWow64 = FALSE;
                                    if (IsWow64Process(hProc, &bIsWow64) && bIsWow64)
                                    {
#ifndef _WIN64
                                        bInject = true;
#endif
                                    }
                                    else
                                    {
#ifdef _WIN64
                                        bInject = true;
#endif
                                    }

                                    if (bInject)
                                    {
                                        if (SUCCEEDED(Inject(hProc, wszModFileName)))
                                        {
                                            std::wcout << std::format(L"Injection succeeded for {} / {}\n", pwszCurrentProc, proc32.th32ProcessID);
                                        }
                                        else
                                        {
                                            std::wcout << std::format(L"Injection failed for {} / {}\n", pwszCurrentProc, proc32.th32ProcessID);
                                        }
                                    }
                                    else
                                    {
                                        std::wcout << std::format(L"Skipping injection on {} / {}. Target arch mismatch\n", pwszCurrentProc, proc32.th32ProcessID);
                                    }
                                }
                                else
                                {
                                    std::wcout << std::format(L"Skipping injection on {} / {}. Already injected\n", pwszCurrentProc, proc32.th32ProcessID);
                                }

                                umapProcsMonitored.insert(std::make_pair(proc32.th32ProcessID, hProc));
                            }
                        }
                    }

                } while (Process32NextW(hProcs, &proc32));

                for (auto iter = umapProcsMonitored.begin(); iter != umapProcsMonitored.end();)
                {
                    DWORD dwExitCode = 0;

                    if (!GetExitCodeProcess(iter->second, &dwExitCode))
                    {
                        std::wcout << std::format(L"GetExitCodeProcess failed on PID {} with {}\n", iter->first, GetLastError());
                    }

                    if (dwExitCode != STILL_ACTIVE)
                    {
                        iter = umapProcsMonitored.erase(iter);
                    }
                    else
                    {
                        ++iter;
                    }
                }
            }
            catch (...)
            {
                ExitPrompt(L"Exception while monitoring for processes\n");
            }
        }
    }
    catch (...)
    {
        ExitPrompt(L"Exception in Run() function\n");
    }
}

extern "C"
{
    __declspec(dllexport) void __cdecl Run()
    {
        Entry(false);
    }
    
    __declspec(dllexport) void __cdecl Run2()
    {
        Entry(true);
    }
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL,
                    DWORD fdwReason,
                    LPVOID lpvReserved)
{
    switch (fdwReason)
    {
    case DLL_PROCESS_ATTACH:
    {
        try
        {
            WCHAR wszModFileName[MAX_PATH] = { 0 };
            PWCHAR pwszExeName = nullptr;

            g_hSelf = hinstDLL;
            g_hklCurrentKeyboard = GetKeyboardLayout(0);

            // Find Russian and English HKLs among installed keyboard layouts
            {
                HKL installedLayouts[64] = {};
                int nInstalled = GetKeyboardLayoutList(_countof(installedLayouts), installedLayouts);
                for (int i = 0; i < nInstalled; i++)
                {
                    LANGID langId = LOWORD(reinterpret_cast<ULONG_PTR>(installedLayouts[i]));
                    switch (PRIMARYLANGID(langId))
                    {
                    case LANG_RUSSIAN: if (!g_hklRussian) g_hklRussian = installedLayouts[i]; break;
                    case LANG_ENGLISH: if (!g_hklEnglish) g_hklEnglish = installedLayouts[i]; break;
                    }
                }
            }

            // Determine initial VM layout state from the VMware thread's keyboard layout
            {
                LANGID langId = LOWORD(reinterpret_cast<ULONG_PTR>(g_hklCurrentKeyboard));
                g_bVMInRussianMode = (PRIMARYLANGID(langId) == LANG_RUSSIAN);
            }

            if (GetModuleFileNameW(nullptr, wszModFileName, _countof(wszModFileName)))
            {
                pwszExeName = wcsrchr(wszModFileName, L'\\');
                pwszExeName = pwszExeName ? pwszExeName + 1 : wszModFileName;

                if (wcscmp(pwszExeName, L"rundll32.exe"))
                {
                    PerformFix();
                }
            }
        }
        catch (...)
        {
            OutputDebugStringW(L"Exception in DLL_PROCESS_ATTACH\n");
            return FALSE;
        }

        break;
    }
    default:
        break;
    }

    return TRUE;
}
