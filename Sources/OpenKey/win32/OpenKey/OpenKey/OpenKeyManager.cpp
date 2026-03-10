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
#include "OpenKeyManager.h"
#include <shlobj.h>

static vector<LPCTSTR> _inputType = {
	_T("Telex"),
	_T("VNI"),
	_T("Simple Telex"),
};

static vector<LPCTSTR> _tableCode = {
	_T("Unicode"),
	_T("TCVN3 (ABC)"),
	_T("VNI Windows"),
	_T("Unicode Tổ hợp"),
	_T("Vietnamese Locale CP 1258")
};

/*-----------------------------------------------------------------------*/

extern void OpenKeyInit();
extern void OpenKeyFree();
extern void ReinstallHooks();
extern void QuickHookCheck();
extern void CheckAndReinstallHooks();
extern void SetSysTrayHwnd(HWND hWnd);
extern void OnForegroundSettled();

unsigned short  OpenKeyManager::_lastKeyCode = 0;

vector<LPCTSTR>& OpenKeyManager::getInputType() {
	return _inputType;
}

vector<LPCTSTR>& OpenKeyManager::getTableCode() {
	return _tableCode;
}

void OpenKeyManager::initEngine() {
	OpenKeyInit();
}

void OpenKeyManager::freeEngine() {
	OpenKeyFree();
}

void OpenKeyManager::reinstallHooks() {
	ReinstallHooks();
}

void OpenKeyManager::quickHookCheck() {
	QuickHookCheck();
}

void OpenKeyManager::checkAndReinstallHooks() {
	CheckAndReinstallHooks();
}

void OpenKeyManager::setSysTrayHwnd(HWND hWnd) {
	SetSysTrayHwnd(hWnd);
}

void OpenKeyManager::onForegroundSettled() {
	OnForegroundSettled();
}

bool OpenKeyManager::checkUpdate(string& newVersion) {
	// Use GitHub Releases API — picks up the latest release automatically whenever
	// the workflow creates a new release, with no need to update version.json manually.
	// API: GET /repos/{owner}/{repo}/releases/latest  → JSON with "tag_name": "v2.0.6"
	wstring dataW = OpenKeyHelper::getContentOfUrl(
		L"https://api.github.com/repos/lotusify/OpenKey/releases/latest");
	string data = wideStringToUtf8(dataW);

	if (data.empty()) return false;

	// Find "tag_name":"v2.0.6"  (GitHub always quotes both key and value)
	const string tagKey = "\"tag_name\"";
	size_t pos = data.find(tagKey);
	if (pos == string::npos) return false;

	// Skip past the key, colon, whitespace, and opening quote
	pos = data.find('\"', pos + tagKey.size()); // opening quote of value
	if (pos == string::npos) return false;
	pos++; // step inside the quote

	// Skip leading 'v' if present (tag format is "v2.0.6")
	if (pos < data.size() && data[pos] == 'v') pos++;

	size_t end = data.find('\"', pos); // closing quote
	if (end == string::npos || end <= pos) return false;

	newVersion = data.substr(pos, end - pos); // e.g. "2.0.6"

	// Parse major.minor.patch from tag string
	int major = 0, minor = 0, patch = 0;
	if (sscanf_s(newVersion.c_str(), "%d.%d.%d", &major, &minor, &patch) != 3)
		return false;

	// Pack into DWORD same way getVersionNumber() does: (major<<16)|(minor<<8)|patch
	DWORD remoteCode = ((DWORD)major << 16) | ((DWORD)minor << 8) | (DWORD)patch;
	DWORD localCode  = OpenKeyHelper::getVersionNumber();

	return remoteCode > localCode;
}

void OpenKeyManager::createDesktopShortcut() {
	CoInitialize(NULL);
	IShellLink* pShellLink = NULL;
	HRESULT hres;
	hres = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_ALL,
						IID_IShellLink, (void**)&pShellLink);
	if (SUCCEEDED(hres)) {
		wstring path = OpenKeyHelper::getFullPath();
		pShellLink->SetPath(path.c_str());
		pShellLink->SetDescription(_T("OpenKey - Bộ gõ Tiếng Việt"));
		pShellLink->SetIconLocation(path.c_str(), 0);

		IPersistFile* pPersistFile;
		hres = pShellLink->QueryInterface(IID_IPersistFile, (void**)&pPersistFile);

		if (SUCCEEDED(hres)) {
			wchar_t desktopPath[MAX_PATH + 1];
			wchar_t savePath[MAX_PATH + 10];
			SHGetFolderPath(NULL, CSIDL_DESKTOP, NULL, 0, desktopPath);
			wsprintf(savePath, _T("%s\\OpenKey.lnk"), desktopPath);
			hres = pPersistFile->Save(savePath, TRUE);
			pPersistFile->Release();
			pShellLink->Release();
			
			// Notify Shell to refresh icon cache - fixes icon not showing on Windows 10
			SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, NULL, NULL);
		}
	}
}
