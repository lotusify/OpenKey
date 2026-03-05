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
#pragma once
#include <windows.h>

// OKLog — lightweight file logger for OpenKey (Win32)
//
// Log file: %APPDATA%\OpenKey\openkey.log
// Rotation:  when file exceeds LOG_MAX_BYTES, renamed to openkey.log.old
// Thread-safe: CRITICAL_SECTION guards all writes
//
// Usage:
//   OKLog::write("HOOK", "keyboard hook installed handle=%p", hKeyboardHook);
//   OKLog::write("HOOK", "hook DEAD — reinstalling");
//   OKLog::write("TOGGLE", "VI mode ON");

namespace OKLog {

    // Initialize logger (create directory, open file).
    // Called once from OpenKeyInit(). Safe to call multiple times.
    void init();

    // Write one log line:
    //   [HH:MM:SS.mmm] [CATEGORY] message
    // category: short tag like "HOOK", "TOGGLE", "IME", "FOREGROUND", etc.
    // format: printf-style format string (ANSI, UTF-8 friendly)
    void write(const char* category, const char* format, ...);

    // Flush and close the log file (called from OpenKeyFree).
    void close();

} // namespace OKLog
