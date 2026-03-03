/*----------------------------------------------------------
OpenKey - The Cross platform Open source Vietnamese Keyboard application.

Copyright (C) 2019 Mai Vu Tuyen
Contact: maivutuyen.91@gmail.com
Github: https://github.com/tuyenvm/OpenKey
Fanpage: https://www.facebook.com/OpenKeyVN

This file is belong to the OpenKey project, Win32 version
which is released under GPL license.
You can fork, modify, improve this program. If you
redistribute your new version, it MUST be open source.
-----------------------------------------------------------*/
#include "ExcludeAppDialog.h"
#include "ExcludeApp.h"
#include "AppDelegate.h"
#include <psapi.h>
#include <commctrl.h>
#include <commdlg.h>
#include <algorithm>
#include <cctype>

#pragma comment(lib, "psapi.lib")
#pragma comment(lib, "comdlg32.lib")

// Helper: UTF-8 std::string -> std::wstring
static std::wstring toWide(const std::string& s) {
    if (s.empty()) return L"";
    int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, NULL, 0);
    std::wstring w(n - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, &w[0], n);
    return w;
}

// Helper: std::wstring -> lowercase UTF-8 std::string
static std::string toLowerUtf8(const std::wstring& w) {
    if (w.empty()) return "";
    int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, NULL, 0, NULL, NULL);
    std::string s(n - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], n, NULL, NULL);
    std::transform(s.begin(), s.end(), s.begin(), ::tolower);
    return s;
}

ExcludeAppDialog::ExcludeAppDialog(const HINSTANCE& hInstance, const int& resourceId)
    : BaseDialog(hInstance, resourceId) {
}

void ExcludeAppDialog::initDialog() {
    SET_DIALOG_ICON(IDI_APP_ICON);

    hListRunning  = GetDlgItem(hDlg, IDC_LIST_RUNNING_PROCS);
    hListExcluded = GetDlgItem(hDlg, IDC_LIST_EXCLUDED_APPS);
    hEditSearch   = GetDlgItem(hDlg, IDC_EDIT_SEARCH_PROC);
    hBtnAdd       = GetDlgItem(hDlg, IDC_BUTTON_ADD_PROC);
    hBtnBrowse    = GetDlgItem(hDlg, IDC_BUTTON_BROWSE_EXE);
    hBtnRemove    = GetDlgItem(hDlg, IDC_BUTTON_REMOVE_EXCLUDED);
    hBtnRefresh   = GetDlgItem(hDlg, IDC_BUTTON_REFRESH_PROCS);

    // Set up ListViews with a single column spanning full width
    RECT rc;
    GetClientRect(hListRunning, &rc);
    LVCOLUMN lvc = {};
    lvc.mask = LVCF_WIDTH;
    lvc.cx = rc.right - rc.left - 4;
    ListView_InsertColumn(hListRunning, 0, &lvc);

    GetClientRect(hListExcluded, &rc);
    lvc.cx = rc.right - rc.left - 4;
    ListView_InsertColumn(hListExcluded, 0, &lvc);

    // Enable full-row select
    ListView_SetExtendedListViewStyle(hListRunning,  LVS_EX_FULLROWSELECT);
    ListView_SetExtendedListViewStyle(hListExcluded, LVS_EX_FULLROWSELECT);

    // Placeholder text for search box
    SendMessageW(hEditSearch, EM_SETCUEBANNER, FALSE, (LPARAM)L"Tìm kiếm...");

    fillRunningProcesses();
    fillExcludedApps();
}

void ExcludeAppDialog::fillRunningProcesses() {
    _runningProcs.clear();

    // EnumProcesses to get all PIDs
    DWORD pids[1024], cbNeeded = 0;
    if (!EnumProcesses(pids, sizeof(pids), &cbNeeded))
        return;

    DWORD count = cbNeeded / sizeof(DWORD);
    for (DWORD i = 0; i < count; i++) {
        if (pids[i] == 0) continue;
        HANDLE hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pids[i]);
        if (!hProc) continue;

        wchar_t path[MAX_PATH] = {};
        if (GetProcessImageFileNameW(hProc, path, MAX_PATH) > 0) {
            // Extract just the filename
            wchar_t* fname = wcsrchr(path, L'\\');
            std::wstring name = fname ? (fname + 1) : path;
            // Lowercase
            std::wstring lower = name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::towlower);
            // Deduplicate
            if (std::find(_runningProcs.begin(), _runningProcs.end(), lower) == _runningProcs.end()) {
                _runningProcs.push_back(lower);
            }
        }
        CloseHandle(hProc);
    }

    std::sort(_runningProcs.begin(), _runningProcs.end());
    filterRunningProcesses(L"");
}

void ExcludeAppDialog::filterRunningProcesses(const std::wstring& filter) {
    ListView_DeleteAllItems(hListRunning);

    std::wstring lowerFilter = filter;
    std::transform(lowerFilter.begin(), lowerFilter.end(), lowerFilter.begin(), ::towlower);

    LVITEMW lvi = {};
    lvi.mask = LVIF_TEXT;
    lvi.iItem = 0;

    for (const auto& name : _runningProcs) {
        if (!lowerFilter.empty() && name.find(lowerFilter) == std::wstring::npos)
            continue;
        lvi.pszText = const_cast<LPWSTR>(name.c_str());
        ListView_InsertItem(hListRunning, &lvi);
        lvi.iItem++;
    }
}

void ExcludeAppDialog::fillExcludedApps() {
    ListView_DeleteAllItems(hListExcluded);

    const auto& apps = getAllExcludedApps();
    std::vector<std::wstring> sorted;
    for (const auto& s : apps)
        sorted.push_back(toWide(s));
    std::sort(sorted.begin(), sorted.end());

    LVITEMW lvi = {};
    lvi.mask = LVIF_TEXT;
    lvi.iItem = 0;
    for (auto& w : sorted) {
        lvi.pszText = const_cast<LPWSTR>(w.c_str());
        ListView_InsertItem(hListExcluded, &lvi);
        lvi.iItem++;
    }
}

void ExcludeAppDialog::onAddProc() {
    int sel = ListView_GetNextItem(hListRunning, -1, LVNI_SELECTED);
    if (sel < 0) return;

    wchar_t buf[MAX_PATH] = {};
    LVITEMW lvi = {};
    lvi.mask = LVIF_TEXT;
    lvi.iItem = sel;
    lvi.pszText = buf;
    lvi.cchTextMax = MAX_PATH;
    ListView_GetItem(hListRunning, &lvi);

    std::string utf8 = toLowerUtf8(buf);
    if (!utf8.empty()) {
        addExcludedApp(utf8);
        fillExcludedApps();
    }
}

void ExcludeAppDialog::onBrowseExe() {
    wchar_t filePath[MAX_PATH] = {};
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = hDlg;
    ofn.lpstrFilter = L"Executable Files (*.exe)\0*.exe\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile   = filePath;
    ofn.nMaxFile    = MAX_PATH;
    ofn.Flags       = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
    ofn.lpstrTitle  = L"Chọn file .exe cần loại trừ";

    if (GetOpenFileNameW(&ofn)) {
        // Extract just filename
        wchar_t* fname = wcsrchr(filePath, L'\\');
        std::wstring name = fname ? (fname + 1) : filePath;
        std::string utf8 = toLowerUtf8(name);
        if (!utf8.empty()) {
            addExcludedApp(utf8);
            fillExcludedApps();
        }
    }
}

void ExcludeAppDialog::onRemoveExcluded() {
    int sel = ListView_GetNextItem(hListExcluded, -1, LVNI_SELECTED);
    if (sel < 0) return;

    wchar_t buf[MAX_PATH] = {};
    LVITEMW lvi = {};
    lvi.mask = LVIF_TEXT;
    lvi.iItem = sel;
    lvi.pszText = buf;
    lvi.cchTextMax = MAX_PATH;
    ListView_GetItem(hListExcluded, &lvi);

    std::string utf8 = toLowerUtf8(buf);
    if (!utf8.empty()) {
        removeExcludedApp(utf8);
        fillExcludedApps();
    }
}

void ExcludeAppDialog::onSearch() {
    wchar_t buf[256] = {};
    GetWindowTextW(hEditSearch, buf, 256);
    filterRunningProcesses(buf);
}

void ExcludeAppDialog::onRefresh() {
    wchar_t buf[256] = {};
    GetWindowTextW(hEditSearch, buf, 256);
    fillRunningProcesses();
    filterRunningProcesses(buf);
}

void ExcludeAppDialog::fillData() {
    fillExcludedApps();
}

INT_PTR ExcludeAppDialog::eventProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_INITDIALOG:
        this->hDlg = hDlg;
        initDialog();
        return TRUE;

    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case IDOK:
        case IDCANCEL:
            AppDelegate::getInstance()->closeDialog(this);
            break;
        case IDC_BUTTON_ADD_PROC:
            onAddProc();
            break;
        case IDC_BUTTON_BROWSE_EXE:
            onBrowseExe();
            break;
        case IDC_BUTTON_REMOVE_EXCLUDED:
            onRemoveExcluded();
            break;
        case IDC_BUTTON_REFRESH_PROCS:
            onRefresh();
            break;
        case IDC_EDIT_SEARCH_PROC:
            if (HIWORD(wParam) == EN_CHANGE)
                onSearch();
            break;
        }
        break;
    }

    case WM_NOTIFY: {
        LPNMHDR pnm = (LPNMHDR)lParam;
        if (pnm->code == NM_DBLCLK && pnm->hwndFrom == hListRunning) {
            onAddProc();
        }
        break;
    }
    }
    return FALSE;
}
