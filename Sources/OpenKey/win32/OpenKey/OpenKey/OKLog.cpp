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
#include "OKLog.h"
#include <windows.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdarg.h>
#include <string>

#pragma comment(lib, "shell32.lib")

// Max log file size before rotation (1 MB)
#define LOG_MAX_BYTES (1024 * 1024)

namespace OKLog {

static CRITICAL_SECTION _cs;
static bool             _csInit    = false;
static FILE*            _fp        = nullptr;
static bool             _inited    = false;
static char             _logPath[MAX_PATH]  = { 0 };
static char             _oldPath[MAX_PATH]  = { 0 };

// Build %APPDATA%\OpenKey\openkey.log path into buf (char, MAX_PATH).
// Returns true on success.
static bool buildLogPath(char* buf, char* oldBuf) {
    wchar_t appDataW[MAX_PATH] = { 0 };
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_APPDATA, NULL, 0, appDataW))) {
        return false;
    }
    // Convert to ANSI (log filenames are ASCII-safe)
    char appData[MAX_PATH] = { 0 };
    WideCharToMultiByte(CP_ACP, 0, appDataW, -1, appData, MAX_PATH, NULL, NULL);

    // Ensure %APPDATA%\OpenKey exists
    char dir[MAX_PATH];
    sprintf_s(dir, MAX_PATH, "%s\\OpenKey", appData);
    CreateDirectoryA(dir, NULL); // OK if already exists

    sprintf_s(buf,    MAX_PATH, "%s\\openkey.log",     dir);
    sprintf_s(oldBuf, MAX_PATH, "%s\\openkey.log.old", dir);
    return true;
}

static void rotateIfNeeded() {
    if (!_fp) return;
    long pos = ftell(_fp);
    if (pos < 0 || pos < LOG_MAX_BYTES) return;

    fclose(_fp);
    _fp = nullptr;

    // Remove old backup, rename current → .old
    DeleteFileA(_oldPath);
    MoveFileA(_logPath, _oldPath);

    fopen_s(&_fp, _logPath, "a");
    if (_fp) {
        fprintf(_fp, "--- log rotated ---\n");
        fflush(_fp);
    }
}

void init() {
    if (_inited) return;

    if (!_csInit) {
        InitializeCriticalSection(&_cs);
        _csInit = true;
    }

    EnterCriticalSection(&_cs);

    if (!_inited) {
        if (buildLogPath(_logPath, _oldPath)) {
            fopen_s(&_fp, _logPath, "a");
            if (_fp) {
                // Write session header
                SYSTEMTIME st;
                GetLocalTime(&st);
                fprintf(_fp,
                    "\n=== OpenKey session start %04d-%02d-%02d %02d:%02d:%02d ===\n",
                    st.wYear, st.wMonth, st.wDay,
                    st.wHour, st.wMinute, st.wSecond);
                fflush(_fp);
            }
        }
        _inited = true;
    }

    LeaveCriticalSection(&_cs);
}

void write(const char* category, const char* format, ...) {
    if (!_inited) init();

    // Format the user message first (outside the lock to reduce contention)
    char msg[1024] = { 0 };
    va_list args;
    va_start(args, format);
    vsnprintf_s(msg, sizeof(msg), _TRUNCATE, format, args);
    va_end(args);

    // Timestamp
    SYSTEMTIME st;
    GetLocalTime(&st);

    EnterCriticalSection(&_cs);

    if (_fp) {
        rotateIfNeeded();
        fprintf(_fp, "[%02d:%02d:%02d.%03d] [%-10s] %s\n",
            st.wHour, st.wMinute, st.wSecond, st.wMilliseconds,
            category, msg);
        fflush(_fp);
    }

    LeaveCriticalSection(&_cs);

    // Also mirror to OutputDebugString for DebugView
    char dbg[1100];
    sprintf_s(dbg, sizeof(dbg), "OpenKey [%s] %s\n", category, msg);
    OutputDebugStringA(dbg);
}

void close() {
    if (!_inited) return;
    EnterCriticalSection(&_cs);
    if (_fp) {
        SYSTEMTIME st;
        GetLocalTime(&st);
        fprintf(_fp,
            "=== OpenKey session end %04d-%02d-%02d %02d:%02d:%02d ===\n",
            st.wYear, st.wMonth, st.wDay,
            st.wHour, st.wMinute, st.wSecond);
        fclose(_fp);
        _fp = nullptr;
    }
    _inited = false;
    LeaveCriticalSection(&_cs);
}

} // namespace OKLog
