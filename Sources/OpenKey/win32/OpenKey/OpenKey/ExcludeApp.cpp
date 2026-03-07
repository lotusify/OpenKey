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
#include "ExcludeApp.h"
#include <windows.h>
#include <sstream>
#include <algorithm>

static const LPCTSTR sk = TEXT("SOFTWARE\\TuyenMai\\OpenKey");
static const LPCTSTR EXCLUDED_APPS_KEY = TEXT("excludedApps");

static std::unordered_set<std::string> _excludedApps;

bool isExcludedApp(const std::string& exeName) {
    std::string lower = exeName;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    return _excludedApps.count(lower) > 0;
}

void addExcludedApp(const std::string& exeName) {
    if (exeName.empty()) return;
    std::string lower = exeName;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    _excludedApps.insert(lower);
    saveExcludedApps();
}

void removeExcludedApp(const std::string& exeName) {
    std::string lower = exeName;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    _excludedApps.erase(lower);
    saveExcludedApps();
}

const std::unordered_set<std::string>& getAllExcludedApps() {
    return _excludedApps;
}

void loadExcludedApps() {
    _excludedApps.clear();

    HKEY hKey = NULL;
    LONG nError = RegOpenKeyEx(HKEY_CURRENT_USER, sk, 0, KEY_READ, &hKey);
    if (nError != ERROR_SUCCESS) return;

    // Query the size of the value first
    DWORD dataSize = 0;
    DWORD type = 0;
    nError = RegQueryValueEx(hKey, EXCLUDED_APPS_KEY, 0, &type, NULL, &dataSize);
    if (nError != ERROR_SUCCESS || dataSize == 0) {
        RegCloseKey(hKey);
        return;
    }

    // Read the value as wide string (REG_SZ)
    std::vector<wchar_t> buf(dataSize / sizeof(wchar_t) + 1, 0);
    nError = RegQueryValueEx(hKey, EXCLUDED_APPS_KEY, 0, &type,
        reinterpret_cast<LPBYTE>(buf.data()), &dataSize);
    RegCloseKey(hKey);

    if (nError != ERROR_SUCCESS) return;

    // Convert to UTF-8
    int utf8Size = WideCharToMultiByte(CP_UTF8, 0, buf.data(), -1, NULL, 0, NULL, NULL);
    if (utf8Size <= 1) return;
    std::string utf8(utf8Size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, buf.data(), -1, &utf8[0], utf8Size, NULL, NULL);

    // Parse pipe-delimited list
    std::istringstream ss(utf8);
    std::string token;
    while (std::getline(ss, token, '|')) {
        if (!token.empty()) {
            _excludedApps.insert(token);
        }
    }
}

void saveExcludedApps() {
    // Build pipe-delimited string
    std::string joined;
    for (const auto& name : _excludedApps) {
        if (!joined.empty()) joined += '|';
        joined += name;
    }

    // Convert to wide string
    int wideSize = MultiByteToWideChar(CP_UTF8, 0, joined.c_str(), -1, NULL, 0);
    std::vector<wchar_t> wide(wideSize, 0);
    MultiByteToWideChar(CP_UTF8, 0, joined.c_str(), -1, wide.data(), wideSize);

    HKEY hKey = NULL;
    LONG nError = RegOpenKeyEx(HKEY_CURRENT_USER, sk, 0, KEY_SET_VALUE, &hKey);
    if (nError == ERROR_FILE_NOT_FOUND) {
        nError = RegCreateKeyEx(HKEY_CURRENT_USER, sk, 0, NULL,
            REG_OPTION_NON_VOLATILE, KEY_CREATE_SUB_KEY | KEY_SET_VALUE, NULL, &hKey, NULL);
    }
    if (nError != ERROR_SUCCESS) return;

    RegSetValueEx(hKey, EXCLUDED_APPS_KEY, 0, REG_SZ,
        reinterpret_cast<const BYTE*>(wide.data()),
        (DWORD)(wide.size() * sizeof(wchar_t)));
    RegCloseKey(hKey);
}
