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
#include "SystemTrayHelper.h"
#include "AppDelegate.h"
#include "OpenKeyManager.h"
#include "OKLog.h"
#include <Wtsapi32.h>

#pragma comment(lib, "Wtsapi32.lib")

#define TIMER_REINSTALL_HOOKS    1001
#define TIMER_HOOK_HEALTH_CHECK  1002   // 1s  — NULL-only check (cheap, no hook removal)
#define TIMER_HOOK_ZOMBIE_CHECK  1003   // 30s — zombie-check via UnhookWindowsHookEx
#define TIMER_FOREGROUND_DEBOUNCE 1004  // 150ms one-shot: fires after foreground settles

// 1s: cheap NULL check — detects hook killed by Windows (handle goes NULL).
// Does NOT call UnhookWindowsHookEx, so it never removes an alive hook.
#define HOOK_HEALTH_CHECK_INTERVAL_MS  1000

// 10s: zombie-check — detects non-NULL but dead handles using UnhookWindowsHookEx.
// Runs infrequently so it won't interrupt typing sessions.
#define HOOK_ZOMBIE_CHECK_INTERVAL_MS  10000

#define WM_TRAYMESSAGE (WM_USER + 1)
#define TRAY_ICONUID 100

#define POPUP_VIET_ON_OFF 900
#define POPUP_SPELLING 901
#define POPUP_SMART_SWITCH 902
#define POPUP_USE_MACRO 903

#define POPUP_TELEX 910
#define POPUP_VNI 911
#define POPUP_SIMPLE_TELEX 912

#define POPUP_UNICODE 930
#define POPUP_TCVN3 931
#define POPUP_VNI_WINDOWS 932
#define POPUP_UNICODE_COMPOUND 933
#define POPUP_VN_LOCALE_1258 934

#define POPUP_CONVERT_TOOL 980
#define POPUP_QUICK_CONVERT 981

#define POPUP_MACRO_TABLE 990

#define POPUP_CONTROL_PANEL 1000
#define POPUP_ABOUT_OPENKEY 1010
#define POPUP_OPENKEY_EXIT 2000

#define MODIFY_MENU(MENU, COMMAND, DATA) ModifyMenu(MENU, COMMAND, \
											MF_BYCOMMAND | (DATA ? MF_CHECKED : MF_UNCHECKED), \
											COMMAND, \
											menuData[COMMAND]);

static HMENU popupMenu;
//static HMENU menuInputType;
static HMENU otherCode;

static NOTIFYICONDATA nid;
static ULONGLONG lastUnlockTime = 0;

#define SESSION_UNLOCK_DEBOUNCE_MS 2000

map<UINT, LPCTSTR> menuData = {
	{POPUP_VIET_ON_OFF, _T("Bật Tiếng Việt")},
	{POPUP_SPELLING, _T("Bật kiểm tra chính tả")},
	{POPUP_SMART_SWITCH, _T("Bật loại trừ ứng dụng thông minh")},
	{POPUP_USE_MACRO, _T("Bật gõ tắt")},
	{POPUP_TELEX, _T("Kiểu gõ Telex")},
	{POPUP_VNI, _T("Kiểu gõ VNI")},
	{POPUP_SIMPLE_TELEX, _T("Kiểu gõ Simple Telex")},
	{POPUP_UNICODE, _T("Unicode dựng sẵn")},
	{POPUP_TCVN3, _T("TCVN3 (ABC)")},
	{POPUP_VNI_WINDOWS, _T("VNI Windows")},
	{POPUP_UNICODE_COMPOUND, _T("Unicode tổ hợp")},
	{POPUP_VN_LOCALE_1258, _T("Vietnamese locale CP 1258")},
	{POPUP_CONVERT_TOOL, _T("Công cụ chuyển mã...")},
	{POPUP_QUICK_CONVERT, _T("Chuyển mã nhanh")},
	{POPUP_MACRO_TABLE, _T("Cấu hình gõ tắt...")},
	{POPUP_CONTROL_PANEL, _T("Bảng điều khiển...")},
	{POPUP_ABOUT_OPENKEY, _T("Giới thiệu OpenKey")},
	{POPUP_OPENKEY_EXIT, _T("Thoát")},
};

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
	static UINT taskbarCreated;

	switch (message) {
	case WM_CREATE:
		taskbarCreated = RegisterWindowMessage(_T("TaskbarCreated"));
		
		// Register session notification for lock/unlock detection
		if (!WTSRegisterSessionNotification(hWnd, NOTIFY_FOR_THIS_SESSION)) {
			OutputDebugString(_T("OpenKey: Failed to register session notification\n"));
		} else {
			OutputDebugString(_T("OpenKey: Session notification registered successfully\n"));
		}

		// Start periodic hook health-check timer (every 30 seconds)
		SetTimer(hWnd, TIMER_HOOK_HEALTH_CHECK, HOOK_HEALTH_CHECK_INTERVAL_MS, NULL);
		SetTimer(hWnd, TIMER_HOOK_ZOMBIE_CHECK,  HOOK_ZOMBIE_CHECK_INTERVAL_MS,  NULL);

		// Pass hWnd to engine so winEventProcCallback can post debounce timers here
		OpenKeyManager::setSysTrayHwnd(hWnd);
		break;
	case WM_USER+2019:
		AppDelegate::getInstance()->onControlPanel();
		break;
	case WM_USER+10: // Mouse click asked main thread to refresh IME state (Fix A)
		OpenKeyManager::refreshImeState();
		break;
		
	// Handle session change (lock/unlock)
	case WM_WTSSESSION_CHANGE:
		if (wParam == WTS_SESSION_LOCK) {
			OKLog::write("SESSION", "screen locked");
		} else if (wParam == WTS_SESSION_UNLOCK) {
			// Debounce: Only process if at least 2 seconds apart
			ULONGLONG now = GetTickCount64();
			if (lastUnlockTime == 0 || now - lastUnlockTime > SESSION_UNLOCK_DEBOUNCE_MS) {
				lastUnlockTime = now;
				OKLog::write("SESSION", "screen unlocked — scheduling hook reinstall in 500ms");
				
				// Use timer to delay 500ms, then reinstall from main thread
				SetTimer(hWnd, TIMER_REINSTALL_HOOKS, 500, NULL);
			}
		}
		break;
		
	// Handle timer for hook reinstallation
	case WM_TIMER:
		if (wParam == TIMER_REINSTALL_HOOKS) {
			KillTimer(hWnd, TIMER_REINSTALL_HOOKS);
			OKLog::write("HOOK", "session-unlock timer fired — reinstalling hooks from main thread");
			OpenKeyManager::reinstallHooks();
		} else if (wParam == TIMER_HOOK_HEALTH_CHECK) {
			// 1s: cheap NULL-only check — never removes an alive hook
			OpenKeyManager::quickHookCheck();
		} else if (wParam == TIMER_HOOK_ZOMBIE_CHECK) {
			// 30s: full zombie-check via UnhookWindowsHookEx
			OpenKeyManager::checkAndReinstallHooks();
		} else if (wParam == TIMER_FOREGROUND_DEBOUNCE) {
			// Foreground has been stable for 150ms — safe to clear session and apply smart-switch
			KillTimer(hWnd, TIMER_FOREGROUND_DEBOUNCE);
			OpenKeyManager::onForegroundSettled();
		}
		break;
	case WM_TRAYMESSAGE: {
		if (lParam == WM_LBUTTONDBLCLK) {
			AppDelegate::getInstance()->onControlPanel();
		}
		if (lParam == WM_LBUTTONUP) {
			AppDelegate::getInstance()->onToggleVietnamese();
			SystemTrayHelper::updateData();
		} else if (lParam == WM_RBUTTONDOWN) {
			POINT curPoint;
			GetCursorPos(&curPoint);
			SetForegroundWindow(hWnd);
			UINT commandId = TrackPopupMenu(
				popupMenu,
				TPM_RETURNCMD | TPM_NONOTIFY,
				curPoint.x,
				curPoint.y,
				0,
				hWnd,
				NULL
			);
			switch (commandId) {
			case POPUP_VIET_ON_OFF:
				AppDelegate::getInstance()->onToggleVietnamese();
				break;
			case POPUP_SPELLING:
				AppDelegate::getInstance()->onToggleCheckSpelling();
				break;
			case POPUP_SMART_SWITCH:
				AppDelegate::getInstance()->onToggleUseSmartSwitchKey();
				break;
			case POPUP_USE_MACRO:
				AppDelegate::getInstance()->onToggleUseMacro();
				break;
			case POPUP_MACRO_TABLE:
				AppDelegate::getInstance()->onMacroTable();
				break;
			case POPUP_CONVERT_TOOL:
				AppDelegate::getInstance()->onConvertTool();
				break;
			case POPUP_QUICK_CONVERT:
				AppDelegate::getInstance()->onQuickConvert();
				break;
			case POPUP_TELEX:
				AppDelegate::getInstance()->onInputType(0);
				break;
			case POPUP_VNI:
				AppDelegate::getInstance()->onInputType(1);
				break;
			case POPUP_SIMPLE_TELEX:
				AppDelegate::getInstance()->onInputType(2);
				break;
			case POPUP_UNICODE:
				AppDelegate::getInstance()->onTableCode(0);
				break;
			case POPUP_TCVN3:
				AppDelegate::getInstance()->onTableCode(1);
				break;
			case POPUP_VNI_WINDOWS:
				AppDelegate::getInstance()->onTableCode(2);
				break;
			case POPUP_UNICODE_COMPOUND:
				AppDelegate::getInstance()->onTableCode(3);
				break;
			case POPUP_VN_LOCALE_1258:
				AppDelegate::getInstance()->onTableCode(4);
				break;
			case POPUP_CONTROL_PANEL:
				AppDelegate::getInstance()->onControlPanel();
				break;
			case POPUP_ABOUT_OPENKEY:
				AppDelegate::getInstance()->onOpenKeyAbout();
				break;
			case POPUP_OPENKEY_EXIT:
				AppDelegate::getInstance()->onOpenKeyExit();
				break;
			}
			SystemTrayHelper::updateData();
		}
	}
	break;
	
	case WM_DESTROY:
		// Kill timers if still active
		KillTimer(hWnd, TIMER_REINSTALL_HOOKS);
		KillTimer(hWnd, TIMER_HOOK_HEALTH_CHECK);
		
		// Unregister session notification on destroy
		WTSUnRegisterSessionNotification(hWnd);
		OutputDebugString(_T("OpenKey: Session notification unregistered\n"));
		break;
		
	default:
		// if the taskbar is restarted, add the system tray icon again
		if (message == taskbarCreated) {
			Shell_NotifyIcon(NIM_ADD, &nid);
		}
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

HWND SystemTrayHelper::createFakeWindow(const HINSTANCE & hIns) {
	//create fake window
	WNDCLASSEXW wcex;
	wcex.cbSize = sizeof(WNDCLASSEX);
	wcex.style = 0;
	wcex.lpfnWndProc = WndProc;
	wcex.cbClsExtra = 0;
	wcex.cbWndExtra = 0;
	wcex.hInstance = hIns;
	wcex.hIcon = LoadIcon(hIns, MAKEINTRESOURCE(IDI_APP_ICON));
	wcex.hCursor = NULL;
	wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wcex.lpszMenuName = NULL;
	wcex.lpszClassName = APP_CLASS;
	wcex.hIconSm = NULL;
	ATOM atom = RegisterClassExW(&wcex);
	HWND hWnd = CreateWindowW(APP_CLASS, _T(""), WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hIns, nullptr);
	if (!hWnd) {
		return NULL;
	}
	ShowWindow(hWnd, 0);
	UpdateWindow(hWnd);
	return hWnd;
}

void SystemTrayHelper::createPopupMenu() {
	popupMenu = CreatePopupMenu();
	AppendMenu(popupMenu, MF_CHECKED, POPUP_VIET_ON_OFF, menuData[POPUP_VIET_ON_OFF]);
	AppendMenu(popupMenu, MF_SEPARATOR, 0, 0);
	AppendMenu(popupMenu, MF_CHECKED, POPUP_SPELLING, menuData[POPUP_SPELLING]);
	AppendMenu(popupMenu, MF_CHECKED, POPUP_SMART_SWITCH, menuData[POPUP_SMART_SWITCH]);
	AppendMenu(popupMenu, MF_CHECKED, POPUP_USE_MACRO, menuData[POPUP_USE_MACRO]);
	AppendMenu(popupMenu, MF_SEPARATOR, 0, 0);
	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_MACRO_TABLE, menuData[POPUP_MACRO_TABLE]);
	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_CONVERT_TOOL, menuData[POPUP_CONVERT_TOOL]);
	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_QUICK_CONVERT, menuData[POPUP_QUICK_CONVERT]);
	AppendMenu(popupMenu, MF_SEPARATOR, 0, 0);

	//menuInputType = CreatePopupMenu();
	AppendMenu(popupMenu, MF_CHECKED, POPUP_TELEX, menuData[POPUP_TELEX]);
	AppendMenu(popupMenu, MF_CHECKED, POPUP_VNI, menuData[POPUP_VNI]);
	AppendMenu(popupMenu, MF_CHECKED, POPUP_SIMPLE_TELEX, menuData[POPUP_SIMPLE_TELEX]);

	//AppendMenu(popupMenu, MF_POPUP, (UINT_PTR)menuInputType, _T("Kiểu gõ"));
	AppendMenu(popupMenu, MF_SEPARATOR, 0, 0);

	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_UNICODE, menuData[POPUP_UNICODE]);
	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_TCVN3, menuData[POPUP_TCVN3]);
	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_VNI_WINDOWS, menuData[POPUP_VNI_WINDOWS]);

	otherCode = CreatePopupMenu();
	AppendMenu(otherCode, MF_CHECKED, POPUP_UNICODE_COMPOUND, menuData[POPUP_UNICODE_COMPOUND]);
	AppendMenu(otherCode, MF_CHECKED, POPUP_VN_LOCALE_1258, menuData[POPUP_VN_LOCALE_1258]);
	AppendMenu(popupMenu, MF_POPUP, (UINT_PTR)otherCode, _T("Bảng mã khác"));

	AppendMenu(popupMenu, MF_SEPARATOR, 0, 0);

	AppendMenu(popupMenu, MF_STRING, POPUP_CONTROL_PANEL, menuData[POPUP_CONTROL_PANEL]);
	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_ABOUT_OPENKEY, menuData[POPUP_ABOUT_OPENKEY]);
	AppendMenu(popupMenu, MF_SEPARATOR, 0, 0);
	AppendMenu(popupMenu, MF_UNCHECKED, POPUP_OPENKEY_EXIT, menuData[POPUP_OPENKEY_EXIT]);

	SetMenuDefaultItem(popupMenu, POPUP_CONTROL_PANEL, false);
}

static void loadTrayIcon() {
	int icon = 0;
	if (vLanguage) {
		if (vUseGrayIcon == 1) icon = IDI_ICON_STATUS_VIET_10;
		else if (vUseGrayIcon == 2) icon = IDI_ICON_STATUS_VIET_10_DARK;
		else icon = IDI_ICON_STATUS_VIET;
		LoadString(GetModuleHandle(0), IDS_TRAY_TITLE_2, nid.szTip, 128);
	}
	else {
		if (vUseGrayIcon == 1) icon = IDI_ICON_STATUS_ENG_10;
		else if (vUseGrayIcon == 2) icon = IDI_ICON_STATUS_ENG_10_DARK;
		else icon = IDI_ICON_STATUS_ENG;
		LoadString(GetModuleHandle(0), IDS_TRAY_TITLE, nid.szTip, 128);
	}
	nid.hIcon = LoadIcon(GetModuleHandle(0), MAKEINTRESOURCE(icon));
}

void SystemTrayHelper::updateData() {
	loadTrayIcon();
	Shell_NotifyIcon(NIM_MODIFY, &nid);

	MODIFY_MENU(popupMenu, POPUP_VIET_ON_OFF, vLanguage);
	MODIFY_MENU(popupMenu, POPUP_SPELLING, vCheckSpelling);
	MODIFY_MENU(popupMenu, POPUP_SMART_SWITCH, vUseSmartSwitchKey);
	MODIFY_MENU(popupMenu, POPUP_USE_MACRO, vUseMacro);
	MODIFY_MENU(popupMenu, POPUP_TELEX, vInputType == 0);
	MODIFY_MENU(popupMenu, POPUP_VNI, vInputType == 1);
	MODIFY_MENU(popupMenu, POPUP_SIMPLE_TELEX, vInputType == 2);
	MODIFY_MENU(popupMenu, POPUP_UNICODE, vCodeTable == 0);
	MODIFY_MENU(popupMenu, POPUP_TCVN3, vCodeTable == 1);
	MODIFY_MENU(popupMenu, POPUP_VNI_WINDOWS, vCodeTable == 2);
	MODIFY_MENU(otherCode, POPUP_UNICODE_COMPOUND, vCodeTable == 3);
	MODIFY_MENU(otherCode, POPUP_VN_LOCALE_1258, vCodeTable == 4);

	wstring hotkey = L"";
	bool hasAdd = false;
	if (convertToolHotKey & 0x100) {
		hotkey += L"Ctrl";
		hasAdd = true;
	}
	if (convertToolHotKey & 0x200) {
		if (hasAdd)
			hotkey += L" + ";
		hotkey += L"Alt";
		hasAdd = true;
	}
	if (convertToolHotKey & 0x400) {
		if (hasAdd)
			hotkey += L" + ";
		hotkey += L"Win";
		hasAdd = true;
	}
	if (convertToolHotKey & 0x800) {
		if (hasAdd)
			hotkey += L" + ";
		hotkey += L"Shift";
		hasAdd = true;
	}

	unsigned short k = ((convertToolHotKey >> 24) & 0xFF);
	if (k != 0xFE) {
		if (hasAdd)
			hotkey += L" + ";
		if (k == VK_SPACE)
			hotkey += L"Space";
		else
			hotkey += (wchar_t)k;
	}

	wstring hotKeyString = menuData[POPUP_QUICK_CONVERT];
	if (hasAdd) {
		hotKeyString += L" - [";
		hotKeyString += hotkey;
		hotKeyString += L"]";
	}
	ModifyMenu(popupMenu, POPUP_QUICK_CONVERT, MF_BYCOMMAND | MF_UNCHECKED, POPUP_QUICK_CONVERT, hotKeyString.c_str());
}

static HINSTANCE ins;
static int recreateCount = 0;

void SystemTrayHelper::_createSystemTrayIcon(const HINSTANCE& hIns) {
	HWND hWnd = createFakeWindow(ins);
	
	if (hWnd == NULL) { //Use timer to create
		if (recreateCount >= 5) {
			PostQuitMessage(0);
			return;
		}
		ins = hIns;
		SetTimer(NULL, 0, 1000 * 3, (TIMERPROC)&WaitToCreateFakeWindow);
		++recreateCount;
		return;
	}
	createPopupMenu();

	//create system tray
	nid.cbSize = sizeof(NOTIFYICONDATA);
	nid.hWnd = hWnd;
	nid.uID = TRAY_ICONUID;
	nid.uVersion = NOTIFYICON_VERSION;
	nid.uCallbackMessage = WM_TRAYMESSAGE;
	loadTrayIcon();
	LoadString(ins, IDS_APP_TITLE, nid.szTip, 128);
	nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;

	// Shell_NotifyIcon may fail if the system tray icon is not fully initialized
	const int maxRetries = 5;
	for (int attempt = 0; attempt < maxRetries; ++attempt) {
		if (Shell_NotifyIcon(NIM_ADD, &nid)) {
			break;
		}
		Sleep(1000);
	}
}


void CALLBACK SystemTrayHelper::WaitToCreateFakeWindow(HWND hwnd, UINT uMsg, UINT timerId, DWORD dwTime) {
	_createSystemTrayIcon(ins);
	KillTimer(0, timerId);
}

void SystemTrayHelper::createSystemTrayIcon(const HINSTANCE& hIns) {
	_createSystemTrayIcon(hIns);
}

void SystemTrayHelper::removeSystemTray() {
	Shell_NotifyIcon(NIM_DELETE, &nid);
}