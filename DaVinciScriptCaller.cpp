#include <windows.h>
#include <tlhelp32.h>
#include <string>
#include <iostream>
#include <vector>
#include <algorithm>
#include <uiautomation.h>
#include <oleauto.h>

static DWORD GetProcessIdByName(const std::wstring& exeName)
{
    DWORD pid = 0;
    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snap == INVALID_HANDLE_VALUE)
        return 0;

    PROCESSENTRY32W pe;
    pe.dwSize = sizeof(pe);
    if (Process32FirstW(snap, &pe))
    {
        do
        {
            if (_wcsicmp(pe.szExeFile, exeName.c_str()) == 0)
            {
                pid = pe.th32ProcessID;
                break;
            }
        } while (Process32NextW(snap, &pe));
    }

    CloseHandle(snap);
    return pid;
}

static BOOL IsMainWindow(HWND hwnd)
{
    return (GetWindow(hwnd, GW_OWNER) == NULL) && IsWindowVisible(hwnd);
}

struct FindWindowForPidContext
{
    DWORD pid;
    HWND hwnd;
};

static BOOL CALLBACK EnumWindowsProc_FindPid(HWND hwnd, LPARAM lParam)
{
    FindWindowForPidContext* ctx = reinterpret_cast<FindWindowForPidContext*>(lParam);
    DWORD wpid = 0;
    GetWindowThreadProcessId(hwnd, &wpid);
    if (wpid == ctx->pid && IsMainWindow(hwnd))
    {
        ctx->hwnd = hwnd;
        return FALSE; // stop enumeration
    }
    return TRUE; // continue
}

static HWND FindMainWindowForPid(DWORD pid)
{
    FindWindowForPidContext ctx;
    ctx.pid = pid;
    ctx.hwnd = NULL;
    EnumWindows(EnumWindowsProc_FindPid, reinterpret_cast<LPARAM>(&ctx));
    return ctx.hwnd;
}

// Fallback: try to find a window whose title contains any of these keywords (case-insensitive).
static HWND FindWindowByTitleKeywords(const std::vector<std::wstring>& keywords)
{
    struct Ctx
    {
        const std::vector<std::wstring>* keywords;
        HWND found;
    } ctx{ &keywords, NULL };

    auto enumProc = [](HWND hwnd, LPARAM lParam) -> BOOL
        {
            Ctx* c = reinterpret_cast<Ctx*>(lParam);
            if (!IsWindowVisible(hwnd))
                return TRUE;

            const int LEN = 512;
            wchar_t buf[LEN] = { 0 };
            if (GetWindowTextW(hwnd, buf, LEN) > 0)
            {
				std::wcout << L"Checking window title: " << buf << L"\n";
                std::wstring title(buf);
                std::wstring ltitle = title;
                std::transform(ltitle.begin(), ltitle.end(), ltitle.begin(), ::towlower);

                for (const auto& kw : *c->keywords)
                {
                    std::wstring lkw = kw;
                    std::transform(lkw.begin(), lkw.end(), lkw.begin(), ::towlower);
					if (ltitle.find(lkw) != std::wstring::npos && ltitle.find(L"script") == std::wstring::npos)
                    {
                        c->found = hwnd;
                        return FALSE; // stop enumeration
                    }
                }
            }
            return TRUE;
        };

    EnumWindows((WNDENUMPROC)enumProc, reinterpret_cast<LPARAM>(&ctx));
    return ctx.found;
}

// Invoke a menu chain like { "Workspace", "Scripts", "Comp", ScripName } using UI Automation.
static bool ClickMenuPath(HWND hwnd, const std::vector<std::wstring>& path)
{
    if (!hwnd)
        return false;

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
        return false;

    IUIAutomation* pAutomation = nullptr;
    hr = CoCreateInstance(CLSID_CUIAutomation, NULL, CLSCTX_INPROC_SERVER,
                          IID_PPV_ARGS(&pAutomation));
    if (FAILED(hr) || !pAutomation)
    {
        CoUninitialize();
        return false;
    }

    IUIAutomationElement* rootElem = nullptr;
    hr = pAutomation->ElementFromHandle(hwnd, &rootElem);
    if (FAILED(hr) || !rootElem)
    {
        pAutomation->Release();
        CoUninitialize();
        return false;
    }

    bool ok = true;
    for (const auto& name : path)
    {
        // Create a condition matching the Name property
        VARIANT var;
        VariantInit(&var);
        var.vt = VT_BSTR;
        var.bstrVal = SysAllocString(name.c_str());

        IUIAutomationCondition* nameCond = nullptr;
        hr = pAutomation->CreatePropertyCondition(UIA_NamePropertyId, var, &nameCond);

        // Free the VARIANT contents we allocated
        VariantClear(&var);

        if (FAILED(hr) || !nameCond)
        {
            ok = false;
            break;
        }

        // Search the whole subtree for the named element (menu items sometimes appear in separate popup trees)
        IUIAutomationElement* found = nullptr;
        hr = rootElem->FindFirst(TreeScope_Descendants, nameCond, &found);
        nameCond->Release();

        if (FAILED(hr) || !found)
        {
            ok = false;
            break;
        }

        // Prefer ExpandCollapse if available, otherwise Invoke
        IUnknown* pPattern = nullptr;
        hr = found->GetCurrentPattern(UIA_ExpandCollapsePatternId, &pPattern);
        if (SUCCEEDED(hr) && pPattern)
        {
            IUIAutomationExpandCollapsePattern* pExp = nullptr;
            hr = pPattern->QueryInterface(IID_PPV_ARGS(&pExp));
            if (SUCCEEDED(hr) && pExp)
            {
                pExp->Expand();
                pExp->Release();
            }
            pPattern->Release();
        }
        else
        {
            pPattern = nullptr;
            hr = found->GetCurrentPattern(UIA_InvokePatternId, &pPattern);
            if (SUCCEEDED(hr) && pPattern)
            {
                IUIAutomationInvokePattern* pInv = nullptr;
                hr = pPattern->QueryInterface(IID_PPV_ARGS(&pInv));
                if (SUCCEEDED(hr) && pInv)
                {
                    pInv->Invoke();
                    pInv->Release();
                }
                pPattern->Release();
            }
        }

        // Allow UI to update/menus to open
        Sleep(100);
        found->Release();
    }

    rootElem->Release();
    pAutomation->Release();
    CoUninitialize();
    return ok;
}

int main(int argc, char* argv[])
{
    if (argc <= 1)
    {
        std::wcout << L"DaVinci Resolve Script Caller\n";
        std::wcout << L"Usage: DaVinciScriptCaller.exe ScriptName\n";
        std::wcout << L"Ensure DaVinci Resolve is running before executing.\n";
        return 0;
	}

    // Try the common executable name for DaVinci Resolve on Windows.
    const std::wstring resolveExeName = L"Resolve.exe";
    DWORD pid = GetProcessIdByName(resolveExeName);

    HWND hwnd = NULL;

    if (pid != 0)
    {
        hwnd = FindMainWindowForPid(pid);
    }

    // Fallback: try to locate a window with "DaVinci" or "Resolve" in its title if exe lookup failed.
    if (pid == 0 || hwnd == NULL)
    {
        std::vector<std::wstring> keywords = { L"DaVinci", L"Resolve" };
        HWND found = FindWindowByTitleKeywords(keywords);
        if (found != NULL)
        {
            DWORD foundPid = 0;
            GetWindowThreadProcessId(found, &foundPid);
            pid = foundPid;
            hwnd = found;
        }
    }

    if (pid != 0)
    {
        std::wcout << L"DaVinci Resolve PID: " << pid << L"\n";
        std::wcout << L"Window handle (HWND): 0x" << std::hex << reinterpret_cast<uintptr_t>(hwnd) << std::dec << L"\n";

        // Convert script name from char* to std::wstring
        int len = static_cast<int>(strlen(argv[1]));
        std::wstring scriptNameW(len, L'\0');
        MultiByteToWideChar(CP_ACP, 0, argv[1], -1, &scriptNameW[0], len + 1);
        scriptNameW.resize(wcslen(scriptNameW.c_str()));

        // Click Workspace->Scripts->Comp->ScriptName
        std::vector<std::wstring> menuPath = { L"Workspace", L"Scripts", L"Comp", scriptNameW };
        bool clicked = ClickMenuPath(hwnd, menuPath);
        if (clicked)
            std::wcout << L"Menu chain invoked.\n";
        else
            std::wcout << L"Failed to invoke menu chain. UI Automation element not found or not invokable.\n";
    }
    else
    {
        std::wcout << L"DaVinci Resolve process/window not found. Ensure Resolve is running.\n";
    }

    return 0;
}

