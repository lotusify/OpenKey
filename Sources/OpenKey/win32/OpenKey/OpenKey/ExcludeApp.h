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
#include "stdafx.h"
#include <unordered_set>

// Excluded app list management.
// Key: exe filename, lowercase UTF-8, e.g. "blender.exe"
// Persistence: HKCU\SOFTWARE\TuyenMai\OpenKey -> "excludedApps" (REG_SZ, pipe-delimited)

bool isExcludedApp(const std::string& exeName);

void addExcludedApp(const std::string& exeName);
void removeExcludedApp(const std::string& exeName);

const std::unordered_set<std::string>& getAllExcludedApps();

void loadExcludedApps();
void saveExcludedApps();
