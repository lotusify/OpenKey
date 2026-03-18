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
#include "stdafx.h"
#include "AppDelegate.h"
#include "ExcludeApp.h"
#include <mutex>

#pragma comment(lib, "imm32")
#define IMC_GETOPENSTATUS 0x0005

#define MASK_SHIFT				0x01
#define MASK_CONTROL			0x02
#define MASK_ALT				0x04
#define MASK_CAPITAL			0x08
#define MASK_NUMLOCK			0x10
#define MASK_WIN				0x20
#define MASK_SCROLL				0x40

#define OTHER_CONTROL_KEY (_flag & MASK_ALT) || (_flag & MASK_CONTROL)
#define DYNA_DATA(macro, pos) (macro ? pData->macroData[pos] : pData->charData[pos])
#define EMPTY_HOTKEY 0xFE0000FE

static vector<string> _chromiumBrowser = {
	"chrome.exe", "brave.exe", "msedge.exe"
};

// MS Office apps that falsely report IME as ON - skip IME check for these
// PowerPoint reports isImeON=1 even when no IME is active, blocking Vietnamese input
static vector<string> _skipImeCheckApps = {
	"powerpnt.exe",   // Microsoft PowerPoint
	"winword.exe",    // Microsoft Word
	"excel.exe"       // Microsoft Excel
};

// Helper for lowercase
static string strToLower(const string& s) {
	string result = s;
	std::transform(result.begin(), result.end(), result.begin(),
		[](unsigned char c) { return std::tolower(c); });
	return result;
}

// Check if current app should skip IME check
// NORMALIZED: compare lowercase to avoid case-sensitivity issues
static bool shouldSkipImeCheck() {
	string lowerAppName = strToLower(OpenKeyHelper::getLastAppExecuteName());
	return std::find(_skipImeCheckApps.begin(), _skipImeCheckApps.end(), lowerAppName) != _skipImeCheckApps.end();
}
extern int vSendKeyStepByStep;
extern int vUseGrayIcon;
extern int vShowOnStartUp;
extern int vRunWithWindows;

static HHOOK hKeyboardHook;
static HHOOK hMouseHook;
static HWINEVENTHOOK hSystemEvent;
static KBDLLHOOKSTRUCT* keyboardData;
static MSLLHOOKSTRUCT* mouseData;
static vKeyHookState* pData;
static vector<Uint16> _syncKey;
static Uint32 _flag = 0, _lastFlag = 0, _privateFlag;
static bool _flagChanged = false, _isFlagKey;
static Uint16 _keycode;
static Uint16 _newChar, _newCharHi;

static vector<Uint16> _newCharString;
static Uint16 _newCharSize;
static bool _willSendControlKey = false;

// Batch input buffer: collects all backspaces + new chars into a single SendInput call.
// Inspired by GoNhanh's TextSender — avoids N separate SendInput calls (one per char)
// and eliminates timing gaps between backspaces and new chars that cause glitches.
static vector<INPUT> _batchInputs;

static Uint16 _uniChar[2];
static int _i, _j, _k;
static Uint32 _tempChar;

static string macroText, macroContent;
static int _languageTemp = 0; //use for smart switch key
static vector<Byte> savedSmartSwitchKeyData; ////use for smart switch key


static bool _hasJustUsedHotKey = false;

// Heartbeat: updated by keyboardHookProcess on every call.
// Main thread health-check reads this to detect zombie hooks (handle non-NULL but hook dead).
// Declared volatile so the compiler doesn't optimize away reads from the main thread.
static volatile ULONGLONG _hookHeartbeat = 0;

// SysTray HWND — set by SystemTrayHelper on WM_CREATE so winEventProcCallback can
// post timers back to the main thread for debouncing foreground changes.
static HWND _sysTrayHwnd = NULL;
void SetSysTrayHwnd(HWND hWnd) { _sysTrayHwnd = hWnd; }

// Timer ID used for debouncing foreground-change session resets.
// Defined here so both winEventProcCallback and SystemTrayHelper WM_TIMER can share it.
#define TIMER_FOREGROUND_DEBOUNCE 1004
#define FOREGROUND_DEBOUNCE_MS       200  // default debounce
#define FOREGROUND_DEBOUNCE_CHROMIUM 500  // Chromium apps fire spurious foreground events

// Chromium-based browsers fire EVENT_SYSTEM_FOREGROUND for internal render/toolbar
// focus changes that don't represent a real app switch. Use a longer debounce so we
// don't call startNewSession() mid-word every time Vivaldi/Chrome blinks its UI.
static const char* _chromiumApps[] = {
	"vivaldi.exe", "chrome.exe", "msedge.exe", "brave.exe",
	"opera.exe", "operagx.exe", "chromium.exe", NULL
};
static bool isChromiumApp(const string& exe) {
	for (int i = 0; _chromiumApps[i]; i++)
		if (_stricmp(exe.c_str(), _chromiumApps[i]) == 0) return true;
	return false;
}

// Magic number to identify OpenKey-generated events (prevent hook re-entry)
// Industry standard practice - UniKey, EVKey use similar approach
#define OPENKEY_EXTRA_INFO 0x4F4B

static INPUT backspaceEvent[2];
static INPUT keyEvent[2];

LRESULT CALLBACK keyboardHookProcess(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK mouseHookProcess(int nCode, WPARAM wParam, LPARAM lParam);
VOID CALLBACK winEventProcCallback(HWINEVENTHOOK hWinEventHook, DWORD dwEvent, HWND hwnd, LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime);
void ReinstallHooks();
void resetForegroundCache(); // forward declare — defined after static cache vars below

void OpenKeyFree() {
	OKLog::write("LIFECYCLE", "OpenKeyFree — shutting down");
	UnhookWindowsHookEx(hMouseHook);
	UnhookWindowsHookEx(hKeyboardHook);
	UnhookWinEvent(hSystemEvent);
	OKLog::close();
}

// Called every 1s from the main thread.
// ONLY checks if the handle is NULL — does NOT call UnhookWindowsHookEx.
// This prevents the "reinstall every second" bug where probing with
// UnhookWindowsHookEx was removing an alive hook, forcing a reinstall each tick.
void QuickHookCheck() {
	if (hKeyboardHook == NULL) {
		OKLog::write("HOOK", "quick-check: keyboard hook handle NULL — reinstalling");
		ReinstallHooks();
		resetForegroundCache();
	}
	// Handle non-NULL → hook is presumed alive; zombie detection handled by 30s timer.
}

// Re-register hooks only — does NOT reset typing state (_flag, _keycode, session).
// Used by zombie-check when hook was alive: we had to remove it to test liveness,
// now we put it back without disturbing an in-progress word.
static void RegisterHooksOnly() {
	static std::mutex registerMutex;
	std::lock_guard<std::mutex> lock(registerMutex);

	HINSTANCE hInstance = GetModuleHandle(NULL);
	hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardHookProcess, hInstance, 0);
	hMouseHook    = SetWindowsHookEx(WH_MOUSE_LL,    mouseHookProcess,    hInstance, 0);

	if (hKeyboardHook && hMouseHook) {
		OKLog::write("HOOK", "RegisterHooksOnly — success (kb=%p, mouse=%p)", hKeyboardHook, hMouseHook);
	} else {
		OKLog::write("HOOK", "RegisterHooksOnly — FAILED (kb=%p, mouse=%p)", hKeyboardHook, hMouseHook);
	}
}

// Called every 10s from the main thread.
// Two-tier liveness check — avoids the "probe by removing" anti-pattern:
//
// Tier 1 (heartbeat): _hookHeartbeat is stamped by keyboardHookProcess on every
//   key event. If it was updated within the last 15s, the hook is definitely alive
//   — no need to touch it at all. This covers the common case (user is typing).
//
// Tier 2 (UnhookWindowsHookEx probe): Only used when the heartbeat is stale (user
//   hasn't typed for >15s). In that case we can't rely on heartbeat alone, so we
//   use the classic UnhookWindowsHookEx probe. If the hook is alive we removed it
//   to test, so we re-register immediately.
//
// This eliminates the 10s periodic unhook→rehook that previously caused ~5ms gaps
// where keystrokes could silently be dropped (no hook present to catch them).
void CheckAndReinstallHooks() {
	if (hKeyboardHook) {
		ULONGLONG now = GetTickCount64();
		ULONGLONG heartbeatAge = now - _hookHeartbeat;

		// Tier 1: heartbeat fresh → hook is alive, no action needed
		if (_hookHeartbeat > 0 && heartbeatAge < 15000) {
			OKLog::write("HOOK", "zombie-check: heartbeat OK (%llums ago) — hook alive, skipping probe", heartbeatAge);
			return; // ← key change: don't touch the hook at all
		}

		// Tier 2: heartbeat stale (user idle or hook may be dead) — probe with UnhookWindowsHookEx
		OKLog::write("HOOK", "zombie-check: heartbeat stale (%llums) — probing with UnhookWindowsHookEx", heartbeatAge);
		if (!UnhookWindowsHookEx(hKeyboardHook)) {
			// Probe failed → hook was already dead (zombie)
			OKLog::write("HOOK", "zombie-check: keyboard hook DEAD — full reinstall");
			hKeyboardHook = NULL;
			if (hMouseHook) { UnhookWindowsHookEx(hMouseHook); hMouseHook = NULL; }
			ReinstallHooks();
			resetForegroundCache();
		} else {
			// Probe succeeded → hook was alive, we just removed it to test — re-register now
			OKLog::write("HOOK", "zombie-check: keyboard hook alive (idle probe) — re-registering only");
			hKeyboardHook = NULL;
			if (hMouseHook) { UnhookWindowsHookEx(hMouseHook); hMouseHook = NULL; }
			RegisterHooksOnly();
		}
	} else {
		OKLog::write("HOOK", "zombie-check: keyboard hook handle NULL — full reinstall");
		if (hMouseHook) { UnhookWindowsHookEx(hMouseHook); hMouseHook = NULL; }
		ReinstallHooks();
		resetForegroundCache();
	}
}


void ReinstallHooks() {
	// Thread-safe: Use static mutex to avoid concurrent reinstalls
	static std::mutex reinstallMutex;
	std::lock_guard<std::mutex> lock(reinstallMutex);
	
	OKLog::write("HOOK", "ReinstallHooks — starting");
	
	// Unhook old hooks (if still active)
	if (hKeyboardHook) {
		UnhookWindowsHookEx(hKeyboardHook);
		hKeyboardHook = NULL;
	}
	
	if (hMouseHook) {
		UnhookWindowsHookEx(hMouseHook);
		hMouseHook = NULL;
	}
	
	// Small delay to ensure hooks are fully released
	Sleep(20);
	
	// Reset state variables
	_lastFlag = 0;
	_keycode = 0;
	_hasJustUsedHotKey = false;
	
	// Resync _flag with current keyboard state
	_flag = 0;
	if (GetKeyState(VK_LSHIFT) < 0 || GetKeyState(VK_RSHIFT) < 0) _flag |= MASK_SHIFT;
	if (GetKeyState(VK_LCONTROL) < 0 || GetKeyState(VK_RCONTROL) < 0) _flag |= MASK_CONTROL;
	if (GetKeyState(VK_LMENU) < 0 || GetKeyState(VK_RMENU) < 0) _flag |= MASK_ALT;
	if (GetKeyState(VK_LWIN) < 0 || GetKeyState(VK_RWIN) < 0) _flag |= MASK_WIN;
	if (GetKeyState(VK_NUMLOCK) < 0) _flag |= MASK_NUMLOCK;
	if (GetKeyState(VK_CAPITAL) & 1) _flag |= MASK_CAPITAL;
	if (GetKeyState(VK_SCROLL) < 0) _flag |= MASK_SCROLL;
	
	// Reinstall hooks
	HINSTANCE hInstance = GetModuleHandle(NULL);
	hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardHookProcess, hInstance, 0);
	hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, mouseHookProcess, hInstance, 0);
	
	if (hKeyboardHook && hMouseHook) {
		OKLog::write("HOOK", "ReinstallHooks — success (kb=%p, mouse=%p)", hKeyboardHook, hMouseHook);
	} else {
		OKLog::write("HOOK", "ReinstallHooks — FAILED (kb=%p, mouse=%p)", hKeyboardHook, hMouseHook);
	}
}

void OpenKeyInit() {
	OKLog::init();
	APP_GET_DATA(vLanguage, 1);
	APP_GET_DATA(vInputType, 0);
	vFreeMark = 0;
	APP_GET_DATA(vCodeTable, 0);
	APP_GET_DATA(vCheckSpelling, 1);
	APP_GET_DATA(vUseModernOrthography, 0);
	APP_GET_DATA(vQuickTelex, 0);
	APP_GET_DATA(vSwitchKeyStatus, 0x7A000206);
	APP_GET_DATA(vRestoreIfWrongSpelling, 1);
	APP_GET_DATA(vFixRecommendBrowser, 1);
	APP_GET_DATA(vUseMacro, 1);
	APP_GET_DATA(vUseMacroInEnglishMode, 0);
	APP_GET_DATA(vAutoCapsMacro, 0);
	APP_GET_DATA(vSendKeyStepByStep, 1);
	APP_GET_DATA(vUseGrayIcon, 0);
	APP_GET_DATA(vShowOnStartUp, 1);
	APP_GET_DATA(vRunWithWindows, 1);
	// #FIXME_UAC: Commented out - causes UAC popup on every startup
	// Call registerRunOnStartup only when user changes setting in UI (see SettingsDialog.cpp, OpenKeySettingsController.cpp)
	// OpenKeyHelper::registerRunOnStartup(vRunWithWindows);
	APP_GET_DATA(vUseSmartSwitchKey, 1);
	APP_GET_DATA(vUpperCaseFirstChar, 0);
	APP_GET_DATA(vAllowConsonantZFWJ, 0);
	APP_GET_DATA(vTempOffSpelling, 0);
	APP_GET_DATA(vQuickStartConsonant, 0);
	APP_GET_DATA(vQuickEndConsonant, 0);
	APP_GET_DATA(vSupportMetroApp, 0);
	APP_GET_DATA(vRunAsAdmin, 0);
	APP_GET_DATA(vCreateDesktopShortcut, 0);
	APP_GET_DATA(vCheckNewVersion, 0);
	APP_GET_DATA(vRememberCode, 1);
	APP_GET_DATA(vOtherLanguage, 1);
	APP_GET_DATA(vTempOffOpenKey, 0);
	APP_GET_DATA(vFixChromiumBrowser, 0);

	//init convert tool
	APP_GET_DATA(convertToolDontAlertWhenCompleted, 0);
	APP_GET_DATA(convertToolToAllCaps, 0);
	APP_GET_DATA(convertToolToAllNonCaps, 0);
	APP_GET_DATA(convertToolToCapsFirstLetter, 0);
	APP_GET_DATA(convertToolToCapsEachWord, 0);
	APP_GET_DATA(convertToolRemoveMark, 0);
	APP_GET_DATA(convertToolFromCode, 0);
	APP_GET_DATA(convertToolToCode, 0);
	APP_GET_DATA(convertToolHotKey, EMPTY_HOTKEY);
	if (convertToolHotKey == 0) {
		convertToolHotKey = EMPTY_HOTKEY;
	}

	pData = (vKeyHookState*)vKeyInit();

	//pre-create back key
	backspaceEvent[0].type = INPUT_KEYBOARD;
	backspaceEvent[0].ki.dwFlags = 0;
	backspaceEvent[0].ki.wVk = VK_BACK;
	backspaceEvent[0].ki.wScan = 0;
	backspaceEvent[0].ki.time = 0;
	backspaceEvent[0].ki.dwExtraInfo = OPENKEY_EXTRA_INFO;

	backspaceEvent[1].type = INPUT_KEYBOARD;
	backspaceEvent[1].ki.dwFlags = KEYEVENTF_KEYUP;
	backspaceEvent[1].ki.wVk = VK_BACK;
	backspaceEvent[1].ki.wScan = 0;
	backspaceEvent[1].ki.time = 0;
	backspaceEvent[1].ki.dwExtraInfo = OPENKEY_EXTRA_INFO;

	//get key state
	_flag = 0;
	if (GetKeyState(VK_LSHIFT) < 0 || GetKeyState(VK_RSHIFT) < 0) _flag |= MASK_SHIFT;
	if (GetKeyState(VK_LCONTROL) < 0 || GetKeyState(VK_RCONTROL) < 0) _flag |= MASK_CONTROL;
	if (GetKeyState(VK_LMENU) < 0 || GetKeyState(VK_RMENU) < 0) _flag |= MASK_ALT;
	if (GetKeyState(VK_LWIN) < 0 || GetKeyState(VK_RWIN) < 0) _flag |= MASK_WIN;
	if (GetKeyState(VK_NUMLOCK) < 0) _flag |= MASK_NUMLOCK;
	if (GetKeyState(VK_CAPITAL) & 1) _flag |= MASK_CAPITAL;
	if (GetKeyState(VK_SCROLL) < 0) _flag |= MASK_SCROLL;

	//init and load macro data
	DWORD macroDataSize;
	BYTE* macroData = OpenKeyHelper::getRegBinary(_T("macroData"), macroDataSize);
	initMacroMap((Byte*)macroData, (int)macroDataSize);

	//init and load smart switch key data
	DWORD smartSwitchKeySize;
	BYTE* data = OpenKeyHelper::getRegBinary(_T("smartSwitchKey"), smartSwitchKeySize);
	initSmartSwitchKey((Byte*)data, (int)smartSwitchKeySize);

	//init and load excluded apps list
	loadExcludedApps();

	//init hook
	HINSTANCE hInstance = GetModuleHandle(NULL);
	hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardHookProcess, hInstance, 0);
	hMouseHook = SetWindowsHookEx(WH_MOUSE_LL, mouseHookProcess, hInstance, 0);
	hSystemEvent = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, NULL, winEventProcCallback, 0, 0, WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);

	OKLog::write("LIFECYCLE", "OpenKeyInit done — lang=%d inputType=%d codeTable=%d spelling=%d kb=%p mouse=%p",
		vLanguage, vInputType, vCodeTable, vCheckSpelling, hKeyboardHook, hMouseHook);
}

void saveSmartSwitchKeyData() {
	getSmartSwitchKeySaveData(savedSmartSwitchKeyData);
	OpenKeyHelper::setRegBinary(_T("smartSwitchKey"), savedSmartSwitchKeyData.data(), (int)savedSmartSwitchKeyData.size());
}

static void InsertKeyLength(const Uint8& len) {
	_syncKey.push_back(len);
}

static inline void prepareKeyEvent(INPUT& input, const Uint16& keycode, const bool& isPress, const DWORD& flag=0) {
	input.type = INPUT_KEYBOARD;
	input.ki.dwFlags = isPress ? flag : flag|KEYEVENTF_KEYUP;
	input.ki.wVk = keycode;
	input.ki.wScan = 0;
	input.ki.time = 0;
	input.ki.dwExtraInfo = OPENKEY_EXTRA_INFO;
}

static inline void prepareUnicodeEvent(INPUT& input, const Uint16& unicode, const bool& isPress) {
	input.type = INPUT_KEYBOARD;
	input.ki.wVk = 0;
	input.ki.wScan = unicode;
	input.ki.time = 0;
	input.ki.dwFlags = (isPress ? 0 : KEYEVENTF_KEYUP) | KEYEVENTF_UNICODE;
	input.ki.dwExtraInfo = OPENKEY_EXTRA_INFO;
}

static void SendCombineKey(const Uint16& key1, const Uint16& key2, const DWORD& flagKey1=0, const DWORD& flagKey2 = 0) {
	prepareKeyEvent(keyEvent[0], key1, true, flagKey1);
	SendInput(1, keyEvent, sizeof(INPUT));

	prepareKeyEvent(keyEvent[0], key2, true, flagKey2);
	prepareKeyEvent(keyEvent[1], key2, false, flagKey2);
	SendInput(2, keyEvent, sizeof(INPUT));

	prepareKeyEvent(keyEvent[0], key1, false, flagKey1);
	SendInput(1, keyEvent, sizeof(INPUT));
}

static void SendKeyCode(Uint32 data) {
	_newChar = (Uint16)data;
	if (!(data & CHAR_CODE_MASK)) {
		if (IS_DOUBLE_CODE(vCodeTable)) //VNI
			InsertKeyLength(1);

		_newChar = keyCodeToCharacter(data);
		if (_newChar == 0) {
			_newChar = (Uint16)data;
			prepareKeyEvent(keyEvent[0], _newChar, true);
			prepareKeyEvent(keyEvent[1], _newChar, false);
			SendInput(2, keyEvent, sizeof(INPUT));
		} else {
			prepareUnicodeEvent(keyEvent[0], _newChar, true);
			prepareUnicodeEvent(keyEvent[1], _newChar, false);
			SendInput(2, keyEvent, sizeof(INPUT));
		}
	} else {
		if (vCodeTable == 0) { //unicode 2 bytes code
			prepareUnicodeEvent(keyEvent[0], _newChar, true);
			prepareUnicodeEvent(keyEvent[1], _newChar, false);
			SendInput(2, keyEvent, sizeof(INPUT));
		} else if (vCodeTable == 1 || vCodeTable == 2 || vCodeTable == 4) { //others such as VNI Windows, TCVN3: 1 byte code
			_newCharHi = HIBYTE(_newChar);
			_newChar = LOBYTE(_newChar);

			prepareUnicodeEvent(keyEvent[0], _newChar, true);
			prepareUnicodeEvent(keyEvent[1], _newChar, false);
			SendInput(2, keyEvent, sizeof(INPUT));

			if (_newCharHi > 32) {
				if (vCodeTable == 2) //VNI
					InsertKeyLength(2);
				prepareUnicodeEvent(keyEvent[0], _newCharHi, true);
				prepareUnicodeEvent(keyEvent[1], _newCharHi, false);
				SendInput(2, keyEvent, sizeof(INPUT));
			} else {
				if (vCodeTable == 2) //VNI
					InsertKeyLength(1);
			}
		} else if (vCodeTable == 3) { //Unicode Compound
			_newCharHi = (_newChar >> 13);
			_newChar &= 0x1FFF;
			_uniChar[0] = _newChar;
			_uniChar[1] = _newCharHi > 0 ? (_unicodeCompoundMark[_newCharHi - 1]) : 0;
			InsertKeyLength(_newCharHi > 0 ? 2 : 1);
			prepareUnicodeEvent(keyEvent[0], _uniChar[0], true);
			prepareUnicodeEvent(keyEvent[1], _uniChar[0], false);
			SendInput(2, keyEvent, sizeof(INPUT));
			if (_newCharHi > 0) {
				prepareUnicodeEvent(keyEvent[0], _uniChar[1], true);
				prepareUnicodeEvent(keyEvent[1], _uniChar[1], false);
				SendInput(2, keyEvent, sizeof(INPUT));
			}
		}
	}
}

static void SendBackspace() {
	SendInput(2, backspaceEvent, sizeof(INPUT));
	if (vSupportMetroApp && OpenKeyHelper::getLastAppExecuteName().compare("ApplicationFrameHost.exe") == 0) {//Metro App
		SendMessage(HWND_BROADCAST, WM_CHAR, VK_BACK, 0L);
		SendMessage(HWND_BROADCAST, WM_CHAR, VK_BACK, 0L);
	}
	if (IS_DOUBLE_CODE(vCodeTable)) { //VNI or Unicode Compound
		if (_syncKey.back() > 1) {
			/*if (!(vCodeTable == 3 && containUnicodeCompoundApp(FRONT_APP))) {
				SendInput(2, backspaceEvent, sizeof(INPUT));
			}*/
			SendInput(2, backspaceEvent, sizeof(INPUT));
			if (vSupportMetroApp && OpenKeyHelper::getLastAppExecuteName().compare("ApplicationFrameHost.exe") == 0) {//Metro App
				SendMessage(HWND_BROADCAST, WM_CHAR, VK_BACK, 0L);
				SendMessage(HWND_BROADCAST, WM_CHAR, VK_BACK, 0L);
			}
		}
		_syncKey.pop_back();
	}
}

// ────────────────────────────────────────────────────────────────────────────
// Batch input helpers — inspired by GoNhanh's TextSender pattern.
// Instead of calling SendInput(2,...) once per character (N round-trips to win32),
// we queue everything into _batchInputs and flush with a single SendInput call.
// This eliminates inter-key timing gaps that cause "lost character" bugs.
// ────────────────────────────────────────────────────────────────────────────

// Queue a backspace into the batch. Preserves VNI/Unicode Compound _syncKey logic.
// Does NOT send to Metro App via SendMessage — Metro App uses the old SendBackspace path.
static void QueueBackspace() {
	_batchInputs.push_back(backspaceEvent[0]);
	_batchInputs.push_back(backspaceEvent[1]);
	if (IS_DOUBLE_CODE(vCodeTable)) { // VNI or Unicode Compound needs extra backspace for double-width chars
		if (_syncKey.back() > 1) {
			_batchInputs.push_back(backspaceEvent[0]);
			_batchInputs.push_back(backspaceEvent[1]);
		}
		_syncKey.pop_back();
	}
}

// Queue a Unicode key event (down or up) into the batch.
static inline void QueueUnicodeEvent(const Uint16& unicode, const bool& isPress) {
	INPUT input;
	input.type = INPUT_KEYBOARD;
	input.ki.wVk = 0;
	input.ki.wScan = unicode;
	input.ki.time = 0;
	input.ki.dwFlags = (isPress ? 0 : KEYEVENTF_KEYUP) | KEYEVENTF_UNICODE;
	input.ki.dwExtraInfo = OPENKEY_EXTRA_INFO;
	_batchInputs.push_back(input);
}

// Queue a virtual-key event (down or up) into the batch.
static inline void QueueVkEvent(const Uint16& vkCode, const bool& isPress, const DWORD& flag = 0) {
	INPUT input;
	input.type = INPUT_KEYBOARD;
	input.ki.dwFlags = isPress ? flag : flag | KEYEVENTF_KEYUP;
	input.ki.wVk = vkCode;
	input.ki.wScan = 0;
	input.ki.time = 0;
	input.ki.dwExtraInfo = OPENKEY_EXTRA_INFO;
	_batchInputs.push_back(input);
}

// Queue a key code into the batch — mirrors SendKeyCode() exactly but queues instead of sending.
static void QueueKeyCode(Uint32 data) {
	_newChar = (Uint16)data;
	if (!(data & CHAR_CODE_MASK)) {
		if (IS_DOUBLE_CODE(vCodeTable)) InsertKeyLength(1);
		_newChar = keyCodeToCharacter(data);
		if (_newChar == 0) {
			_newChar = (Uint16)data;
			QueueVkEvent(_newChar, true);
			QueueVkEvent(_newChar, false);
		} else {
			QueueUnicodeEvent(_newChar, true);
			QueueUnicodeEvent(_newChar, false);
		}
	} else {
		if (vCodeTable == 0) { // Unicode (NFC)
			QueueUnicodeEvent(_newChar, true);
			QueueUnicodeEvent(_newChar, false);
		} else if (vCodeTable == 1 || vCodeTable == 2 || vCodeTable == 4) { // TCVN3, VNI Windows, CP1258
			_newCharHi = HIBYTE(_newChar);
			_newChar   = LOBYTE(_newChar);
			QueueUnicodeEvent(_newChar, true);
			QueueUnicodeEvent(_newChar, false);
			if (_newCharHi > 32) {
				if (vCodeTable == 2) InsertKeyLength(2);
				QueueUnicodeEvent(_newCharHi, true);
				QueueUnicodeEvent(_newCharHi, false);
			} else {
				if (vCodeTable == 2) InsertKeyLength(1);
			}
		} else if (vCodeTable == 3) { // Unicode Compound (NFD)
			_newCharHi = (_newChar >> 13);
			_newChar  &= 0x1FFF;
			_uniChar[0] = _newChar;
			_uniChar[1] = _newCharHi > 0 ? (_unicodeCompoundMark[_newCharHi - 1]) : 0;
			InsertKeyLength(_newCharHi > 0 ? 2 : 1);
			QueueUnicodeEvent(_uniChar[0], true);
			QueueUnicodeEvent(_uniChar[0], false);
			if (_newCharHi > 0) {
				QueueUnicodeEvent(_uniChar[1], true);
				QueueUnicodeEvent(_uniChar[1], false);
			}
		}
	}
}

// Flush all queued inputs with a single SendInput call, then clear the queue.
static void FlushBatchInputs() {
	if (!_batchInputs.empty()) {
		SendInput((UINT)_batchInputs.size(), _batchInputs.data(), sizeof(INPUT));
		_batchInputs.clear();
	}
}



static void SendEmptyCharacter() {
	if (IS_DOUBLE_CODE(vCodeTable)) //VNI or Unicode Compound
		InsertKeyLength(1);

	_newChar = 0x202F; //empty char

	prepareUnicodeEvent(keyEvent[0], _newChar, true);
	prepareUnicodeEvent(keyEvent[1], _newChar, false);
	SendInput(2, keyEvent, sizeof(INPUT));
}

static void SendNewCharString(const bool& dataFromMacro = false) {
	_j = 0;
	_newCharSize = dataFromMacro ? (Uint16)pData->macroData.size() : pData->newCharCount;
	if (_newCharString.size() < _newCharSize) {
		_newCharString.resize(_newCharSize);
	}
	_willSendControlKey = false;
	
	if (_newCharSize > 0) {
		for (_k = dataFromMacro ? 0 : pData->newCharCount - 1;
			dataFromMacro ? _k < pData->macroData.size() : _k >= 0;
			dataFromMacro ? _k++ : _k--) {

			_tempChar = DYNA_DATA(dataFromMacro, _k);
			if (_tempChar & PURE_CHARACTER_MASK) {
				_newCharString[_j++] = _tempChar;
				if (IS_DOUBLE_CODE(vCodeTable)) {
					InsertKeyLength(1);
				}
			} else if (!(_tempChar & CHAR_CODE_MASK)) {
				if (IS_DOUBLE_CODE(vCodeTable)) //VNI
					InsertKeyLength(1);
				_newCharString[_j++] = keyCodeToCharacter(_tempChar);
			} else {
				_newChar = _tempChar;
				if (vCodeTable == 0) {  //unicode 2 bytes code
					_newCharString[_j++] = _newChar;
				} else if (vCodeTable == 1 || vCodeTable == 2 || vCodeTable == 4) { //others such as VNI Windows, TCVN3: 1 byte code
					_newCharHi = HIBYTE(_newChar);
					_newChar = LOBYTE(_newChar);
					_newCharString[_j++] = _newChar;

					if (_newCharHi > 32) {
						if (vCodeTable == 2) //VNI
							InsertKeyLength(2);
						_newCharString[_j++] = _newCharHi;
						_newCharSize++;
					}
					else {
						if (vCodeTable == 2) //VNI
							InsertKeyLength(1);
					}
				} else if (vCodeTable == 3) { //Unicode Compound
					_newCharHi = (_newChar >> 13);
					_newChar &= 0x1FFF;

					InsertKeyLength(_newCharHi > 0 ? 2 : 1);
					_newCharString[_j++] = _newChar;
					if (_newCharHi > 0) {
						_newCharSize++;
						_newCharString[_j++] = _unicodeCompoundMark[_newCharHi - 1];
					}

				}
			}
		}//end for
	}

	if (pData->code == vRestore || pData->code == vRestoreAndStartNewSession) { //if is restore
		if (keyCodeToCharacter(_keycode) != 0) {
			_newCharSize++;
			_newCharString[_j++] = keyCodeToCharacter(_keycode | ((_flag & MASK_SHIFT) || (_flag & MASK_CAPITAL) ? CAPS_MASK : 0));
		} else {
			_willSendControlKey = true;
		}
	}
	if (pData->code == vRestoreAndStartNewSession) {
		startNewSession();
	}

	OpenKeyHelper::setClipboardText((LPCTSTR)_newCharString.data(), _newCharSize + 1, CF_UNICODETEXT);

	// Small yield: allow Windows clipboard subsystem to process the SetClipboardData()
	// before the target app receives Shift+Insert. Without this, fast apps (Chromium, etc.)
	// can read an empty/stale clipboard and paste nothing — causing the "lost character" bug.
	Sleep(2);

	//Send shift + insert
	SendCombineKey(KEY_LEFT_SHIFT, VK_INSERT, 0, KEYEVENTF_EXTENDEDKEY);
	
	//the case when hCode is vRestore or vRestoreAndStartNewSession,
	//the word is invalid and last key is control key such as TAB, LEFT ARROW, RIGHT ARROW,...
	if (_willSendControlKey) {
		SendKeyCode(_keycode);
	}
}

bool checkHotKey(int hotKeyData, bool checkKeyCode = true) {
	if ((hotKeyData & (~0x8000)) == EMPTY_HOTKEY)
		return false;
	if (HAS_CONTROL(hotKeyData) ^ GET_BOOL(_lastFlag & MASK_CONTROL))
		return false;
	if (HAS_OPTION(hotKeyData) ^ GET_BOOL(_lastFlag & MASK_ALT))
		return false;
	if (HAS_COMMAND(hotKeyData) ^ GET_BOOL(_lastFlag & MASK_WIN))
		return false;
	if (HAS_SHIFT(hotKeyData) ^ GET_BOOL(_lastFlag & MASK_SHIFT))
		return false;
	if (checkKeyCode) {
		if (GET_SWITCH_KEY(hotKeyData) != _keycode)
			return false;
	}
	return true;
}

void switchLanguage() {
	int prevLang = vLanguage;
	if (vLanguage == 0)
		vLanguage = 1;
	else
		vLanguage = 0;
	OKLog::write("TOGGLE", "switchLanguage %s -> %s (app=%s flag=0x%02X)",
		prevLang ? "VI" : "EN",
		vLanguage ? "VI" : "EN",
		OpenKeyHelper::getLastAppExecuteName().c_str(), _flag);
	if (HAS_BEEP(vSwitchKeyStatus))
		MessageBeep(MB_OK);
	AppDelegate::getInstance()->onInputMethodChangedFromHotKey();
	if (vUseSmartSwitchKey) {
		setAppInputMethodStatus(OpenKeyHelper::getFrontMostAppExecuteName(), vLanguage | (vCodeTable << 1));
		saveSmartSwitchKeyData();
	}
	OKLog::write("SESSION", "startNewSession — switchLanguage");
	startNewSession();
}

static void SendPureCharacter(const Uint16& ch) {
	if (ch < 128)
		SendKeyCode(ch);
	else {
		prepareUnicodeEvent(keyEvent[0], ch, true);
		prepareUnicodeEvent(keyEvent[1], ch, false);
		SendInput(2, keyEvent, sizeof(INPUT));
		if (IS_DOUBLE_CODE(vCodeTable)) {
			InsertKeyLength(1);
		}
	}
}

static void handleMacro() {
	//fix autocomplete
	if (vFixRecommendBrowser) {
		SendEmptyCharacter();
		pData->backspaceCount++;
	}

	//send backspace
	if (pData->backspaceCount > 0) {
		for (int i = 0; i < pData->backspaceCount; i++) {
			SendBackspace();
		}
	}
	//send real data
	if (!vSendKeyStepByStep) {
		SendNewCharString(true);
	} else {
		for (int i = 0; i < pData->macroData.size(); i++) {
			if (pData->macroData[i] & PURE_CHARACTER_MASK) {
				SendPureCharacter(pData->macroData[i]);
			} else {
				SendKeyCode(pData->macroData[i]);
			}
		}
	}
	SendKeyCode(_keycode | (_flag & MASK_SHIFT ? CAPS_MASK : 0));
}

static bool SetModifierMask(const Uint16& vkCode) {
	// For caps lock case, toggling the flag isn't enough. We need to check the actual state, which should be done before each key press.
	// Example: the caps lock state can be changed without the key being pressed, or the key toggle is made with admin privilege, making the app not able to detect the change.
	Uint32 prevCapsBit = _flag & MASK_CAPITAL;
	if (GetKeyState(VK_CAPITAL) & 1) _flag |= MASK_CAPITAL;
	else _flag &= ~MASK_CAPITAL;

	// Log only when Caps Lock state actually changes
	if ((_flag & MASK_CAPITAL) != prevCapsBit) {
		OKLog::write("INPUT", "CapsLock -> %s (GetKeyState=0x%04X vk=0x%02X)",
			(_flag & MASK_CAPITAL) ? "ON" : "OFF",
			(unsigned)GetKeyState(VK_CAPITAL), vkCode);
	}

	if (vkCode == VK_LSHIFT || vkCode == VK_RSHIFT) _flag |= MASK_SHIFT;
	else if (vkCode == VK_LCONTROL || vkCode == VK_RCONTROL) _flag |= MASK_CONTROL;
	else if (vkCode == VK_LMENU || vkCode == VK_RMENU) _flag |= MASK_ALT;
	else if (vkCode == VK_LWIN || vkCode == VK_RWIN) _flag |= MASK_WIN;
	else if (vkCode == VK_NUMLOCK) _flag |= MASK_NUMLOCK;
	else if (vkCode == VK_SCROLL) _flag |= MASK_SCROLL;
	else { 
		_isFlagKey = false;
		return false; 
	}
	_isFlagKey = true;
	return true;
}

static bool UnsetModifierMask(const Uint16& vkCode) {
	if (vkCode == VK_LSHIFT || vkCode == VK_RSHIFT) _flag &= ~MASK_SHIFT;
	else if (vkCode == VK_LCONTROL || vkCode == VK_RCONTROL) _flag &= ~MASK_CONTROL;
	else if (vkCode == VK_LMENU || vkCode == VK_RMENU) _flag &= ~MASK_ALT;
	else if (vkCode == VK_LWIN || vkCode == VK_RWIN) _flag &= ~MASK_WIN;
	else if (vkCode == VK_NUMLOCK) _flag &= ~MASK_NUMLOCK;
	else if (vkCode == VK_SCROLL) _flag &= ~MASK_SCROLL;
	else { 
		_isFlagKey = false;
		return false; 
	}
	_isFlagKey = true;
	return true;
}

// IME check cache — updated on the MAIN THREAD only (in OnForegroundSettled + mouse handler).
//
// KEY DESIGN: Previously, SendMessageTimeout was called inside keyboardHookProcess (hook callback).
// WH_KEYBOARD_LL callbacks have a hard ~300ms budget; any blocking call risks Windows
// silently killing the hook. Even SMTO_ABORTIFHUNG at 10ms is risky under system load.
//
// New strategy (Fix A):
//   - _cachedImeState is set ONLY from the main thread where there is no timeout budget.
//   - Hook callback is read-only: just tests _cachedImeState (atomic bool, ~nanoseconds).
//   - IME check runs in OnForegroundSettled() which already fires on the main thread.
//   - Mouse click also triggers a fresh check via resetImeSessionCache() + refreshImeState().
//
// This eliminates the last Win32-blocking call from the hot keypress path.
static volatile bool _cachedImeState = false; // written by main thread, read by hook thread

static HWND _cachedForegroundHwnd = NULL; // Cached GetForegroundWindow() result
static HWND _cachedDefaultImeHwnd = NULL; // Cached ImmGetDefaultIMEWnd() result

// Perform the IME state check — MUST only be called from the main thread.
// Reads the current IME open status via SendMessageTimeout and updates _cachedImeState.
void refreshImeState() {
	HWND hFg = GetForegroundWindow();
	HWND hIME = ImmGetDefaultIMEWnd(hFg);
	_cachedForegroundHwnd = hFg;
	_cachedDefaultImeHwnd = hIME;

	bool imeOn = false;
	if (hIME != NULL) {
		DWORD_PTR dwResult = 0;
		// No timeout risk: running on main thread, not inside hook callback.
		if (SendMessageTimeout(hIME, WM_IME_CONTROL, IMC_GETOPENSTATUS, 0,
		                       SMTO_ABORTIFHUNG | SMTO_BLOCK, 50, &dwResult)) {
			imeOn = (dwResult != 0);
		}
	}
	_cachedImeState = imeOn;
	OKLog::write("IME", "refreshImeState: hFg=%p hIME=%p ime=%s app=%s",
		hFg, hIME, imeOn ? "ON" : "OFF",
		OpenKeyHelper::getLastAppExecuteName().c_str());
}

// Reset all foreground-dependent caches (called on foreground window change)
inline void resetForegroundCache() {
	_cachedImeState = false; // safe default until refreshImeState() runs on main thread
	_cachedForegroundHwnd = NULL;
	_cachedDefaultImeHwnd = NULL;
	OpenKeyHelper::invalidateAppNameCache();
}

// Kept for compatibility (mouse click path calls this)
inline void resetImeSessionCache() {
	resetForegroundCache();
}

LRESULT CALLBACK keyboardHookProcess(int nCode, WPARAM wParam, LPARAM lParam) {
	// Update heartbeat so main-thread health-check can detect zombie hooks
	_hookHeartbeat = GetTickCount64();

	keyboardData = (KBDLLHOOKSTRUCT *)lParam;
	//ignore my event (check for OpenKey magic number)
	if (keyboardData->dwExtraInfo == OPENKEY_EXTRA_INFO) {
		return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
	}

	//ignore events from excluded apps (hard exclude - no processing at all)
	if (isExcludedApp(OpenKeyHelper::getLastAppExecuteName())) {
		return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
	}
	
	// IME check: _cachedImeState is maintained by the MAIN THREAD (OnForegroundSettled, mouse).
	// Hook callback is read-only here — no blocking Win32 calls inside the hook.
	HWND hWnd = _cachedForegroundHwnd ? _cachedForegroundHwnd : GetForegroundWindow();
	
	// Only block input when OpenKey is in Vietnamese mode AND IME is genuinely active.
	// When in English mode (vLanguage==0), always pass through — there's no reason to block.
	// Skip IME check for apps that falsely report IME as ON (MS Office, etc.).
	if (vLanguage == 1 && _cachedImeState && !shouldSkipImeCheck()) {
		// Log once per foreground session — avoid flooding on every keystroke
		static HWND _lastImeBlockLoggedHwnd = NULL;
		if (hWnd != _lastImeBlockLoggedHwnd) {
			OKLog::write("INPUT", "IME active — blocking OpenKey input (hwnd=%p app=%s)",
				hWnd, OpenKeyHelper::getLastAppExecuteName().c_str());
			_lastImeBlockLoggedHwnd = hWnd;
		}
		return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
	}
	
	//check modifier key
	if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
		//LOG(L"Key down: %d\n", keyboardData->vkCode);
		SetModifierMask((Uint16)keyboardData->vkCode);
	} else if (wParam == WM_KEYUP || wParam == WM_SYSKEYUP) {
		//LOG(L"Key up: %d\n", keyboardData->vkCode);
		UnsetModifierMask((Uint16)keyboardData->vkCode);
	}
	if (!_isFlagKey && wParam != WM_KEYUP && wParam != WM_SYSKEYUP)
		_keycode = (Uint16)keyboardData->vkCode;

	//switch language shortcut; convert hotkey
	if ((wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) && !_isFlagKey && _keycode != 0) {
		if (GET_SWITCH_KEY(vSwitchKeyStatus) != _keycode && GET_SWITCH_KEY(convertToolHotKey) != _keycode) {
			_lastFlag = 0;
		} else {
			if (GET_SWITCH_KEY(vSwitchKeyStatus) == _keycode && checkHotKey(vSwitchKeyStatus, GET_SWITCH_KEY(vSwitchKeyStatus) != 0xFE)) {
				switchLanguage();
				_hasJustUsedHotKey = true;
				_keycode = 0;
				return -1;
			}
			if (GET_SWITCH_KEY(convertToolHotKey) == _keycode && checkHotKey(convertToolHotKey, GET_SWITCH_KEY(convertToolHotKey) != 0xFE)) {
				AppDelegate::getInstance()->onQuickConvert();
				_hasJustUsedHotKey = true;
				_keycode = 0;
				return -1;
			}
		}
		_hasJustUsedHotKey = _lastFlag != 0;
	} else if (_isFlagKey) {
		if (_lastFlag == 0 || _lastFlag < _flag)
			_lastFlag = _flag;
		else if (_lastFlag > _flag) {
			//check switch
			if (checkHotKey(vSwitchKeyStatus, GET_SWITCH_KEY(vSwitchKeyStatus) != 0xFE)) {
				switchLanguage();
				_hasJustUsedHotKey = true;
			}
			if (checkHotKey(convertToolHotKey, GET_SWITCH_KEY(convertToolHotKey) != 0xFE)) {
				AppDelegate::getInstance()->onQuickConvert();
				_hasJustUsedHotKey = true;
			}
			//check temporarily turn off spell checking
			if (vTempOffSpelling && !_hasJustUsedHotKey && _lastFlag & MASK_CONTROL) {
				vTempOffSpellChecking();
			}
			if (vTempOffOpenKey && !_hasJustUsedHotKey && _lastFlag & MASK_ALT) {
				vTempOffEngine();
			}
			_lastFlag = _flag;
			_hasJustUsedHotKey = false;
		}
		_keycode = 0;
		return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
	}

	//if is in english mode
	if (vLanguage == 0) {
		if (vUseMacro && vUseMacroInEnglishMode && (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN)) {
			vEnglishMode(((wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) ? vKeyEventState::KeyDown : vKeyEventState::MouseDown),
				_keycode,
				(_flag & MASK_SHIFT) || (_flag & MASK_CAPITAL),
				OTHER_CONTROL_KEY);

			if (pData->code == vReplaceMaro) { //handle macro in english mode
				handleMacro();
				return NULL;
			}
		}
		return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
	}

	//handle keyboard
	if (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN) {
		//send event signal to Engine
		vKeyHandleEvent(vKeyEvent::Keyboard,
						vKeyEventState::KeyDown,
						_keycode,
						(_flag & MASK_SHIFT && _flag & MASK_CAPITAL) ? 0 : (_flag & MASK_SHIFT ? 1 : (_flag & MASK_CAPITAL ? 2 : 0)),
						OTHER_CONTROL_KEY);
		if (pData->code == vDoNothing) { //do nothing
			if (IS_DOUBLE_CODE(vCodeTable)) { //VNI
				if (pData->extCode == 1) { //break key
					_syncKey.clear();
				} else if (pData->extCode == 2) { //delete key
					if (_syncKey.size() > 0) {
						if (_syncKey.back() > 1 && (vCodeTable == 2 || vCodeTable == 3)) {
							//send one more backspace
							SendInput(2, backspaceEvent, sizeof(INPUT));
						}
						_syncKey.pop_back();
					}
				} else if (pData->extCode == 3) { //normal key
					InsertKeyLength(1);
				}
			}
			return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
		} else if (pData->code == vWillProcess || pData->code == vRestore || pData->code == vRestoreAndStartNewSession) { //handle result signal
			if (pData->code == vRestore || pData->code == vRestoreAndStartNewSession) {
				OKLog::write("SESSION", "engine restore (code=%d bs=%d chars=%d key=0x%02X app=%s)",
					pData->code, pData->backspaceCount, pData->newCharCount, _keycode,
					OpenKeyHelper::getLastAppExecuteName().c_str());
			}
			//fix autocomplete
			if (vFixRecommendBrowser && pData->extCode != 4) {
			if (vFixChromiumBrowser && 
					std::find(_chromiumBrowser.begin(), _chromiumBrowser.end(), strToLower(OpenKeyHelper::getLastAppExecuteName())) != _chromiumBrowser.end()) {
					SendCombineKey(KEY_LEFT_SHIFT, KEY_LEFT, 0, KEYEVENTF_EXTENDEDKEY);
					if (pData->backspaceCount == 1)
						pData->backspaceCount--;
				} else {
					SendEmptyCharacter();
					pData->backspaceCount++;
				}
			}
			
			//send backspace + new character
			// Batch strategy (inspired by GoNhanh's TextSender):
			//   - For normal apps: queue backspaces + chars → single SendInput call.
			//     This eliminates timing gaps between individual SendInput(2,...) calls
			//     that can cause apps to miss characters at high typing speed.
			//   - Metro App: must use old sequential path (needs SendMessage in between).
			//   - Clipboard path (!vSendKeyStepByStep): keeps existing behavior.
			if (!vSendKeyStepByStep) {
				// Clipboard path — unchanged (handles complex char data via Shift+Insert)
				if (pData->backspaceCount > 0 && pData->backspaceCount < MAX_BUFF) {
					for (_i = 0; _i < pData->backspaceCount; _i++) {
						SendBackspace();
					}
				}
				SendNewCharString();
			} else if (vSupportMetroApp && OpenKeyHelper::getLastAppExecuteName().compare("ApplicationFrameHost.exe") == 0) {
				// Metro App: sequential path (SendMessage cannot be batched in SendInput)
				if (pData->backspaceCount > 0 && pData->backspaceCount < MAX_BUFF) {
					for (_i = 0; _i < pData->backspaceCount; _i++) {
						SendBackspace();
					}
				}
				if (pData->newCharCount > 0 && pData->newCharCount <= MAX_BUFF) {
					for (int i = pData->newCharCount - 1; i >= 0; i--) {
						SendKeyCode(pData->charData[i]);
					}
				}
				if (pData->code == vRestore || pData->code == vRestoreAndStartNewSession) {
					SendKeyCode(_keycode | ((_flag & MASK_CAPITAL) || (_flag & MASK_SHIFT) ? CAPS_MASK : 0));
				}
				if (pData->code == vRestoreAndStartNewSession) {
					OKLog::write("SESSION", "startNewSession — vRestoreAndStartNewSession (key=0x%02X)", _keycode);
					startNewSession();
				}
			} else {
				// ★ Batch path — default for all normal apps (no Metro App, no clipboard)
				// Build entire replacement operation as one INPUT array:
				//   [backspace×N] + [char×M] → single SendInput call
				_batchInputs.clear();
				_batchInputs.reserve((pData->backspaceCount + pData->newCharCount + 1) * 4);

				// Queue backspaces (preserves VNI _syncKey logic inside QueueBackspace)
				if (pData->backspaceCount > 0 && pData->backspaceCount < MAX_BUFF) {
					for (_i = 0; _i < pData->backspaceCount; _i++) {
						QueueBackspace();
					}
				}

				// Queue new characters
				if (pData->newCharCount > 0 && pData->newCharCount <= MAX_BUFF) {
					for (int i = pData->newCharCount - 1; i >= 0; i--) {
						QueueKeyCode(pData->charData[i]);
					}
				}

				// Queue restore key (when word is invalid, re-emit the triggering key)
				if (pData->code == vRestore || pData->code == vRestoreAndStartNewSession) {
					QueueKeyCode(_keycode | ((_flag & MASK_CAPITAL) || (_flag & MASK_SHIFT) ? CAPS_MASK : 0));
				}

				// ONE SendInput call for everything — no race conditions, no timing gaps
				FlushBatchInputs();

				if (pData->code == vRestoreAndStartNewSession) {
					OKLog::write("SESSION", "startNewSession — vRestoreAndStartNewSession (key=0x%02X)", _keycode);
					startNewSession();
				}
			}
		} else if (pData->code == vReplaceMaro) { //MACRO
			handleMacro();
		}
		return -1; //consume event
	}
	return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}

LRESULT CALLBACK mouseHookProcess(int nCode, WPARAM wParam, LPARAM lParam) {
	mouseData = (MSLLHOOKSTRUCT *)lParam;
	switch (wParam) {
	case WM_LBUTTONDOWN:
		// Reset IME cache on left click — user may have changed focus.
		// Post to main thread so refreshImeState() runs there (no hook-timeout risk).
		resetForegroundCache(); // immediate invalidate (cheap)
		if (_sysTrayHwnd)
			PostMessage(_sysTrayHwnd, WM_USER + 10, 0, 0); // asks main thread to call refreshImeState()
		// fall through
	
	case WM_RBUTTONDOWN:
	case WM_MBUTTONDOWN:
	case WM_XBUTTONDOWN:
	case WM_NCXBUTTONDOWN:
	case WM_LBUTTONUP:
	case WM_RBUTTONUP:
	case WM_MBUTTONUP:
	case WM_XBUTTONUP:
	case WM_NCXBUTTONUP:
		//send event signal to Engine
		vKeyHandleEvent(vKeyEvent::Mouse, vKeyEventState::MouseDown, 0);
		if (IS_DOUBLE_CODE(vCodeTable)) { //VNI
			_syncKey.clear();
		}
		break;
	}
	return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

// Called from SystemTrayHelper WM_TIMER after FOREGROUND_DEBOUNCE_MS of stable foreground.
// Contains the original foreground-change logic (smart switch, session reset).
// Separated from winEventProcCallback to allow debouncing.
void OnForegroundSettled() {
	string& exe = OpenKeyHelper::getFrontMostAppExecuteName();
	OKLog::write("FOREGROUND", "app settled -> %s", exe.c_str());

	// Refresh IME state now that we're on the main thread.
	// This is the ONLY place where SendMessageTimeout is called — safe, no hook timeout risk.
	if (!shouldSkipImeCheck()) {
		refreshImeState();
	} else {
		_cachedImeState = false; // skip-IME apps: always treat as IME off
	}
	// Excluded app check runs unconditionally — independent of smart switch setting.
	// When switching to an excluded app, force EN mode and clear session.
	if (isExcludedApp(exe)) {
		OKLog::write("FOREGROUND", "excluded app — forcing EN");
		if (vLanguage != 0) {
			vLanguage = 0;
			AppDelegate::getInstance()->onInputMethodChangedFromHotKey();
		}
		OKLog::write("SESSION", "startNewSession — excluded app (%s)", exe.c_str());
		startNewSession();
		return;
	}

	if (vUseSmartSwitchKey || vRememberCode) {
		if (exe.compare("explorer.exe") == 0) //dont apply with windows explorer
			return;
		_languageTemp = getAppInputMethodStatus(exe, vLanguage | (vCodeTable << 1));
		vTempOffEngine(false);
		if (vUseSmartSwitchKey && (_languageTemp & 0x01) != vLanguage) {
			if (_languageTemp != -1) {
				OKLog::write("TOGGLE", "smartSwitch %s -> %s for app=%s",
					vLanguage ? "VI" : "EN",
					(_languageTemp & 0x01) ? "VI" : "EN",
					exe.c_str());
				vLanguage = _languageTemp;
				AppDelegate::getInstance()->onInputMethodChangedFromHotKey();
			} else {
				saveSmartSwitchKeyData();
			}
		}
		OKLog::write("SESSION", "startNewSession — foreground settled (%s lang=%s)", exe.c_str(), vLanguage ? "VI" : "EN");
		startNewSession();
		if (vRememberCode && (_languageTemp >> 1) != vCodeTable) {
			if (_languageTemp != -1) {
				AppDelegate::getInstance()->onTableCode(_languageTemp >> 1);
			} else {
				saveSmartSwitchKeyData();
			}
		}
		if (vSupportMetroApp && exe.compare("ApplicationFrameHost.exe") == 0) {
			SendMessage(HWND_BROADCAST, WM_CHAR, VK_BACK, 0L);
			SendMessage(HWND_BROADCAST, WM_CHAR, VK_BACK, 0L);
		}
	} else {
		OKLog::write("SESSION", "startNewSession — foreground settled (%s lang=%s)", exe.c_str(), vLanguage ? "VI" : "EN");
		startNewSession();
	}
}

VOID CALLBACK winEventProcCallback(HWINEVENTHOOK hWinEventHook, DWORD dwEvent, HWND hwnd, LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime) {
	// Foreground window changed — invalidate caches immediately (cheap, always safe).
	resetForegroundCache();

	// Fix B: Chromium apps (Vivaldi, Chrome, Edge...) fire EVENT_SYSTEM_FOREGROUND for
	// internal render/tab focus changes that don't represent a real app switch.
	// Use a longer debounce so we don't nuke the typing session mid-word.
	if (_sysTrayHwnd) {
		string& exe = OpenKeyHelper::getFrontMostAppExecuteName();
		UINT debounceMs = isChromiumApp(exe) ? FOREGROUND_DEBOUNCE_CHROMIUM : FOREGROUND_DEBOUNCE_MS;
		KillTimer(_sysTrayHwnd, TIMER_FOREGROUND_DEBOUNCE);
		SetTimer(_sysTrayHwnd, TIMER_FOREGROUND_DEBOUNCE, debounceMs, NULL);
	}
}
