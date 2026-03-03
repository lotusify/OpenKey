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
#include "BaseDialog.h"
#include "resource.h"
#include <vector>
#include <string>

class ExcludeAppDialog : public BaseDialog {
private:
    HWND hListRunning = NULL;
    HWND hListExcluded = NULL;
    HWND hEditSearch = NULL;
    HWND hBtnAdd = NULL;
    HWND hBtnBrowse = NULL;
    HWND hBtnRemove = NULL;
    HWND hBtnRefresh = NULL;

    // Full list of running processes (exe names, lowercase)
    std::vector<std::wstring> _runningProcs;

    void initDialog();
    void fillRunningProcesses();
    void filterRunningProcesses(const std::wstring& filter);
    void fillExcludedApps();
    void onAddProc();
    void onBrowseExe();
    void onRemoveExcluded();
    void onSearch();
    void onRefresh();

protected:
    virtual INT_PTR eventProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam) override;

public:
    ExcludeAppDialog(const HINSTANCE& hInstance, const int& resourceId);
    virtual void fillData() override;
};
