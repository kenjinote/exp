#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <windowsx.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <shellapi.h>
#include <pathcch.h>
#include <strsafe.h>
#include "resource.h"
#include <vector>
#include <string>
#include <algorithm>
#include <shlobj.h>
#include <memory>
#include <dwmapi.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <Aclapi.h> // ★追加: 所有者情報取得のため

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "pathcch.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "advapi32.lib") // ★追加: 所有者情報取得のため

#define IDC_TAB 1000
#define EDIT_ID_BASE 1001
#define LIST_ID_BASE 2001
#define IDC_ADD_TAB_BUTTON 3001
#define SEARCH_TIMER_ID 1

// SDKのバージョンによっては定義されていないため、手動で定義する
#ifndef TCBS_NORMAL
#define TCBS_NORMAL 1
#define TCBS_HOT 2
#endif

#ifndef TABP_CLOSEBUTTON
#define TABP_CLOSEBUTTON 8
#endif

// キャプションボタンの非アクティブ状態
#ifndef MINBS_INACTIVE
#define MINBS_INACTIVE 5
#endif
#ifndef MAXBS_INACTIVE
#define MAXBS_INACTIVE 5
#endif
#ifndef RBS_INACTIVE
#define RBS_INACTIVE 5
#endif
#ifndef CBS_INACTIVE
#define CBS_INACTIVE 5
#endif

// キャプションボタンの押下状態
#ifndef MINBS_PUSHED
#define MINBS_PUSHED 3
#endif
#ifndef MAXBS_PUSHED
#define MAXBS_PUSHED 3
#endif
#ifndef RBS_PUSHED
#define RBS_PUSHED 3
#endif
#ifndef CBS_PUSHED
#define CBS_PUSHED 3
#endif

// ★追加: ファイル情報と所有者名を保持する構造体
struct FileInfo {
    WIN32_FIND_DATAW findData;
    std::wstring owner;
};

struct ExplorerTabData {
    HWND hEdit;
    HWND hList;
    WCHAR currentPath[MAX_PATH];
    std::vector<FileInfo> fileData; // ★修正: WIN32_FIND_DATAW から変更
    int sortColumn;
    bool sortAscending;
    std::wstring searchString;
    UINT_PTR searchTimerId;
    bool bProgrammaticPathChange;
};

WCHAR szClassName[] = L"exp";
static HWND hTab;
static HWND hAddButton;
static HFONT hFont;
static HIMAGELIST hImgList;
static std::vector<std::unique_ptr<ExplorerTabData>> g_tabs;
static int g_hoveredCloseButtonTab = -1;

static RECT g_rcMinButton, g_rcMaxButton, g_rcCloseButton;
enum class CaptionButtonState { None, Min, Max, Close };
static CaptionButtonState g_hoveredButton = CaptionButtonState::None;
static CaptionButtonState g_pressedButton = CaptionButtonState::None;
static bool g_isTrackingMouse = false;

void SwitchTab(HWND hWnd, int newIndex);
void UpdateLayout(HWND hWnd);
void AddNewTab(HWND hWnd, LPCWSTR initialPath);
void CloseTab(HWND hWnd, int index);
LRESULT CALLBACK EditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK ListSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
void ListDirectory(ExplorerTabData* pData);
void UpdateSortMark(ExplorerTabData* pData);
void UpdateTabTitle(ExplorerTabData* pData);
void UpdateCaptionButtonsRect(HWND hWnd);
void DrawCaptionButtons(HDC hdc, HWND hWnd);


ExplorerTabData* GetCurrentTabData() {
    int sel = TabCtrl_GetCurSel(hTab);
    if (sel >= 0 && sel < (int)g_tabs.size()) {
        return g_tabs[sel].get();
    }
    return nullptr;
}

ExplorerTabData* GetTabDataFromChild(HWND hChild) {
    for (const auto& tab : g_tabs) {
        if (tab->hEdit == hChild || tab->hList == hChild) {
            return tab.get();
        }
    }
    return nullptr;
}

// ★追加: ファイル所有者を取得するヘルパー関数
std::wstring GetFileOwnerW(const WCHAR* filePath) {
    PSID pSidOwner = NULL;
    PSECURITY_DESCRIPTOR pSD = NULL;
    std::wstring ownerName;

    DWORD dwRtn = GetNamedSecurityInfoW(
        filePath,
        SE_FILE_OBJECT,
        OWNER_SECURITY_INFORMATION,
        &pSidOwner,
        NULL,
        NULL,
        NULL,
        (LPVOID*)&pSD);

    if (dwRtn == ERROR_SUCCESS) {
        WCHAR szAccountName[256];
        WCHAR szDomainName[256];
        DWORD dwAccountNameSize = ARRAYSIZE(szAccountName);
        DWORD dwDomainNameSize = ARRAYSIZE(szDomainName);
        SID_NAME_USE eUse;

        if (LookupAccountSidW(
            NULL,
            pSidOwner,
            szAccountName,
            &dwAccountNameSize,
            szDomainName,
            &dwDomainNameSize,
            &eUse)) {
            ownerName = szAccountName;
        }
        else {
            // アカウント名の解決に失敗した場合
            ownerName = L"N/A";
        }
        LocalFree(pSD);
    }
    else {
        // GetNamedSecurityInfoに失敗した場合
        ownerName = L"";
    }

    return ownerName;
}

bool CompareFunction(const FileInfo& a, const FileInfo& b, ExplorerTabData* pData) { // ★修正: 引数の型を変更
    if (wcscmp(a.findData.cFileName, L"..") == 0) return true;
    if (wcscmp(b.findData.cFileName, L"..") == 0) return false;
    const bool isDirA = (a.findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
    const bool isDirB = (b.findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
    if (isDirA != isDirB) {
        return isDirA > isDirB;
    }
    int result = 0;
    switch (pData->sortColumn) {
    case 0:
        result = StrCmpLogicalW(a.findData.cFileName, b.findData.cFileName);
        break;
    case 1:
    {
        ULARGE_INTEGER sizeA = { a.findData.nFileSizeLow, a.findData.nFileSizeHigh };
        ULARGE_INTEGER sizeB = { b.findData.nFileSizeLow, b.findData.nFileSizeHigh };
        if (sizeA.QuadPart < sizeB.QuadPart) result = -1;
        else if (sizeA.QuadPart > sizeB.QuadPart) result = 1;
        else result = 0;
        break;
    }
    case 2:
        result = CompareFileTime(&a.findData.ftLastWriteTime, &b.findData.ftLastWriteTime);
        break;
    case 3:
        result = CompareFileTime(&a.findData.ftCreationTime, &b.findData.ftCreationTime);
        break;
    case 4: // ★追加: 所有者でのソート
        result = a.owner.compare(b.owner);
        break;
    }
    return pData->sortAscending ? (result < 0) : (result > 0);
}

void ListDirectory(ExplorerTabData* pData) {
    pData->fileData.clear();
    if (PathIsRootW(pData->currentPath) == FALSE) {
        FileInfo up_info = { 0 }; // ★修正
        up_info.findData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        StringCchCopyW(up_info.findData.cFileName, MAX_PATH, L"..");
        pData->fileData.push_back(up_info);
    }
    WIN32_FIND_DATAW fd;
    WCHAR searchPath[MAX_PATH];
    PathCombineW(searchPath, pData->currentPath, L"*");
    HANDLE hFind = FindFirstFileW(searchPath, &fd);
    if (hFind != INVALID_HANDLE_VALUE) {
        do {
            if (wcscmp(fd.cFileName, L".") != 0 && wcscmp(fd.cFileName, L"..") != 0) {
                FileInfo info; // ★追加
                info.findData = fd; // ★追加

                // ★追加: 所有者情報を取得
                WCHAR fullPath[MAX_PATH];
                PathCombineW(fullPath, pData->currentPath, fd.cFileName);
                info.owner = GetFileOwnerW(fullPath);

                pData->fileData.push_back(info); // ★修正
            }
        } while (FindNextFileW(hFind, &fd));
        FindClose(hFind);
    }
    std::sort(pData->fileData.begin(), pData->fileData.end(), [&](const FileInfo& a, const FileInfo& b) { // ★修正
        return CompareFunction(a, b, pData);
        });
    ListView_SetItemCountEx(pData->hList, pData->fileData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(pData->hList, NULL, TRUE);
}

void GetSelectedFilePathsW(ExplorerTabData* pData, std::vector<std::wstring>& paths) {
    paths.clear();
    int iItem = -1;
    while ((iItem = ListView_GetNextItem(pData->hList, iItem, LVNI_SELECTED)) != -1) {
        if (iItem < (int)pData->fileData.size()) {
            const WIN32_FIND_DATAW& fd = pData->fileData[iItem].findData; // ★修正
            if (wcscmp(fd.cFileName, L"..") != 0) {
                WCHAR fullPath[MAX_PATH];
                PathCombineW(fullPath, pData->currentPath, fd.cFileName);
                paths.push_back(fullPath);
            }
        }
    }
}

void DoCopyCut(HWND hWnd, bool isCut) {
    ExplorerTabData* pData = GetCurrentTabData();
    if (!pData) return;
    std::vector<std::wstring> paths;
    GetSelectedFilePathsW(pData, paths);
    if (paths.empty()) return;
    size_t totalLen = 0;
    for (const auto& s : paths) {
        totalLen += s.length() + 1;
    }
    totalLen++;
    HGLOBAL hGlobal = GlobalAlloc(GHND | GMEM_SHARE, sizeof(DROPFILES) + totalLen * sizeof(WCHAR));
    if (!hGlobal) return;
    DROPFILES* df = (DROPFILES*)GlobalLock(hGlobal);
    if (!df) {
        GlobalFree(hGlobal);
        return;
    }
    df->pFiles = sizeof(DROPFILES);
    df->fWide = TRUE;
    WCHAR* buffer = (WCHAR*)((BYTE*)df + sizeof(DROPFILES));
    WCHAR* current = buffer;
    for (const auto& s : paths) {
        StringCchCopyW(current, totalLen - (current - buffer), s.c_str());
        current += s.length() + 1;
    }
    *current = L'\0';
    GlobalUnlock(hGlobal);
    if (OpenClipboard(hWnd)) {
        EmptyClipboard();
        SetClipboardData(CF_HDROP, hGlobal);
        if (isCut) {
            HGLOBAL hEffect = GlobalAlloc(GHND | GMEM_SHARE, sizeof(DWORD));
            if (hEffect) {
                DWORD* pEffect = (DWORD*)GlobalLock(hEffect);
                if (pEffect) {
                    *pEffect = DROPEFFECT_MOVE;
                    GlobalUnlock(hEffect);
                    SetClipboardData(RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT), hEffect);
                }
            }
        }
        CloseClipboard();
    }
    else {
        GlobalFree(hGlobal);
    }
}

void DoPaste(HWND hWnd) {
    ExplorerTabData* pData = GetCurrentTabData();
    if (!pData) return;
    if (!IsClipboardFormatAvailable(CF_HDROP)) return;
    if (!OpenClipboard(hWnd)) return;
    HDROP hDrop = (HDROP)GetClipboardData(CF_HDROP);
    if (hDrop) {
        bool isCut = false;
        HGLOBAL hEffect = GetClipboardData(RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT));
        if (hEffect) {
            DWORD* pEffect = (DWORD*)GlobalLock(hEffect);
            if (pEffect && (*pEffect & DROPEFFECT_MOVE)) {
                isCut = true;
            }
            GlobalUnlock(hEffect);
        }
        UINT fileCount = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
        if (fileCount > 0) {
            size_t fromPathLen = 0;
            for (UINT i = 0; i < fileCount; ++i) {
                fromPathLen += DragQueryFileW(hDrop, i, NULL, 0) + 1;
            }
            fromPathLen++;
            std::vector<WCHAR> fromPaths(fromPathLen, 0);
            WCHAR* current = fromPaths.data();
            for (UINT i = 0; i < fileCount; i++) {
                current += DragQueryFileW(hDrop, i, current, fromPathLen - (current - fromPaths.data()));
                current++;
            }
            SHFILEOPSTRUCTW sfo = { 0 };
            sfo.hwnd = hWnd;
            sfo.wFunc = isCut ? FO_MOVE : FO_COPY;
            sfo.pFrom = fromPaths.data();
            sfo.pTo = pData->currentPath;
            sfo.fFlags = FOF_ALLOWUNDO;
            CloseClipboard();
            int result = SHFileOperationW(&sfo);
            if (result == 0 && isCut) {
                if (OpenClipboard(NULL)) {
                    EmptyClipboard();
                    CloseClipboard();
                }
            }
        }
        else {
            CloseClipboard();
        }
    }
    else {
        CloseClipboard();
    }
    ListDirectory(pData);
}

void DoDelete(HWND hWnd, bool permanent) {
    ExplorerTabData* pData = GetCurrentTabData();
    if (!pData) return;
    std::vector<std::wstring> paths;
    GetSelectedFilePathsW(pData, paths);
    if (paths.empty()) return;
    size_t totalLen = 0;
    for (const auto& s : paths) {
        totalLen += s.length() + 1;
    }
    totalLen++;
    std::vector<WCHAR> fromPaths(totalLen, 0);
    WCHAR* current = fromPaths.data();
    for (const auto& s : paths) {
        StringCchCopyW(current, totalLen - (current - fromPaths.data()), s.c_str());
        current += s.length() + 1;
    }
    SHFILEOPSTRUCTW sfo = { 0 };
    sfo.hwnd = hWnd;
    sfo.wFunc = FO_DELETE;
    sfo.pFrom = fromPaths.data();
    sfo.fFlags = permanent ? 0 : FOF_ALLOWUNDO;
    SHFileOperationW(&sfo);
    ListDirectory(pData);
}

void UpdateSortMark(ExplorerTabData* pData) {
    HWND hHeader = ListView_GetHeader(pData->hList);
    for (int i = 0; i < Header_GetItemCount(hHeader); ++i) {
        HDITEM hdi = { HDI_FORMAT };
        Header_GetItem(hHeader, i, &hdi);
        hdi.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
        Header_SetItem(hHeader, i, &hdi);
    }
    HDITEM hdi = { HDI_FORMAT };
    Header_GetItem(hHeader, pData->sortColumn, &hdi);
    hdi.fmt |= (pData->sortAscending ? HDF_SORTUP : HDF_SORTDOWN);
    Header_SetItem(hHeader, pData->sortColumn, &hdi);
}

void FileTimeToString(const FILETIME& ft, WCHAR* buf, size_t bufSize) {
    SYSTEMTIME stUTC, stLocal;
    FileTimeToSystemTime(&ft, &stUTC);
    SystemTimeToTzSpecificLocalTime(NULL, &stUTC, &stLocal);
    StringCchPrintf(buf, bufSize, L"%04d/%02d/%02d %02d:%02d",
        stLocal.wYear, stLocal.wMonth, stLocal.wDay, stLocal.wHour, stLocal.wMinute);
}

void ShowContextMenu(HWND hWnd, ExplorerTabData* pData, int iItem, POINT pt) {
    if (iItem < -1) return;
    HRESULT hr;
    IShellFolder* pDesktopFolder = nullptr;
    IShellFolder* pParentFolder = nullptr;
    LPITEMIDLIST pidlParent = nullptr;
    IContextMenu* pContextMenu = nullptr;
    std::vector<LPCITEMIDLIST> selectedPidls;

    if (FAILED(SHGetDesktopFolder(&pDesktopFolder))) goto cleanup;
    if (FAILED(pDesktopFolder->ParseDisplayName(hWnd, NULL, pData->currentPath, NULL, &pidlParent, NULL))) goto cleanup;
    if (FAILED(pDesktopFolder->BindToObject(pidlParent, NULL, IID_IShellFolder, (void**)&pParentFolder))) goto cleanup;

    if (iItem != -1) {
        int currentSel = -1;
        while ((currentSel = ListView_GetNextItem(pData->hList, currentSel, LVNI_SELECTED)) != -1) {
            if (currentSel < (int)pData->fileData.size()) {
                const WIN32_FIND_DATAW& fd = pData->fileData[currentSel].findData; // ★修正
                if (wcscmp(fd.cFileName, L"..") == 0) continue;
                LPITEMIDLIST pidlItem = nullptr;
                if (SUCCEEDED(pParentFolder->ParseDisplayName(hWnd, NULL, (LPWSTR)fd.cFileName, NULL, &pidlItem, NULL))) {
                    selectedPidls.push_back(pidlItem);
                }
            }
        }
    }
    if (!selectedPidls.empty()) {
        hr = pParentFolder->GetUIObjectOf(hWnd, selectedPidls.size(), selectedPidls.data(), IID_IContextMenu, NULL, (void**)&pContextMenu);
    }
    else {
        hr = pParentFolder->CreateViewObject(hWnd, IID_IContextMenu, (void**)&pContextMenu);
    }

    if (SUCCEEDED(hr)) {
        HMENU hMenu = CreatePopupMenu();
        if (hMenu) {
            if (SUCCEEDED(pContextMenu->QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL))) {
                UINT idCmd = TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
                if (idCmd > 0) {
                    CMINVOKECOMMANDINFO cmi = { sizeof(cmi) };
                    cmi.hwnd = hWnd;
                    cmi.lpVerb = MAKEINTRESOURCEA(idCmd - 1);
                    cmi.nShow = SW_SHOWNORMAL;
                    pContextMenu->InvokeCommand(&cmi);
                    ListDirectory(pData);
                }
            }
            DestroyMenu(hMenu);
        }
    }

cleanup:
    if (pContextMenu) pContextMenu->Release();
    for (auto& pidl : selectedPidls) CoTaskMemFree((LPVOID)pidl);
    if (pParentFolder) pParentFolder->Release();
    if (pidlParent) CoTaskMemFree(pidlParent);
    if (pDesktopFolder) pDesktopFolder->Release();
}

void NavigateUpAndSelect(HWND hWnd) {
    ExplorerTabData* pData = GetCurrentTabData();
    if (!pData) return;

    WCHAR previousFolderName[MAX_PATH];
    if (PathIsRootW(pData->currentPath) == FALSE) {
        StringCchCopyW(previousFolderName, MAX_PATH, PathFindFileNameW(pData->currentPath));
        PathCchRemoveFileSpec(pData->currentPath, MAX_PATH);

        UpdateTabTitle(pData);

        SetWindowTextW(pData->hEdit, pData->currentPath);
        ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED);
        ListDirectory(pData);
        UpdateSortMark(pData);
        for (size_t i = 0; i < pData->fileData.size(); ++i) {
            if (wcscmp(pData->fileData[i].findData.cFileName, previousFolderName) == 0) { // ★修正
                ListView_SetItemState(pData->hList, (int)i, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
                ListView_EnsureVisible(pData->hList, (int)i, FALSE);
                break;
            }
        }
    }
}

void NavigateToSelected(ExplorerTabData* pData) {
    int sel = ListView_GetNextItem(pData->hList, -1, LVNI_SELECTED);
    if (sel < 0 || sel >= (int)pData->fileData.size()) {
        return;
    }
    const WIN32_FIND_DATAW& fd = pData->fileData[sel].findData; // ★修正
    if (wcscmp(fd.cFileName, L"..") == 0) {
        NavigateUpAndSelect(GetParent(pData->hList));
        return;
    }
    if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
        WCHAR newPath[MAX_PATH];
        PathCombineW(newPath, pData->currentPath, fd.cFileName);
        StringCchCopyW(pData->currentPath, MAX_PATH, newPath);

        UpdateTabTitle(pData);

        SetWindowTextW(pData->hEdit, pData->currentPath);
        ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED);
        ListDirectory(pData);
        UpdateSortMark(pData);
        if (ListView_GetItemCount(pData->hList) > 0) {
            ListView_SetItemState(pData->hList, 0, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
            ListView_EnsureVisible(pData->hList, 0, FALSE);
        }
    }
    else {
        WCHAR fullpath[MAX_PATH];
        PathCombineW(fullpath, pData->currentPath, fd.cFileName);
        ShellExecuteW(NULL, L"open", fullpath, NULL, NULL, SW_SHOWNORMAL);
    }
}

void UpdateTabTitle(ExplorerTabData* pData) {
    int tabIndex = -1;
    for (size_t i = 0; i < g_tabs.size(); ++i) {
        if (g_tabs[i].get() == pData) {
            tabIndex = (int)i;
            break;
        }
    }
    if (tabIndex == -1) return;

    TCITEMW tci = { TCIF_TEXT };
    WCHAR tabText[MAX_PATH];
    StringCchCopyW(tabText, MAX_PATH, PathFindFileNameW(pData->currentPath));
    if (wcslen(tabText) == 0) {
        StringCchCopyW(tabText, MAX_PATH, pData->currentPath);
        PathRemoveBackslashW(tabText);
    }
    tci.pszText = tabText;

    TabCtrl_SetItem(hTab, tabIndex, &tci);

    UpdateLayout(GetParent(hTab));
}

LRESULT CALLBACK EditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    ExplorerTabData* pData = GetTabDataFromChild(hWnd);
    if (!pData) return DefSubclassProc(hWnd, uMsg, wParam, lParam);

    switch (uMsg)
    {
    case WM_KEYDOWN:
        if (wParam == VK_RETURN)
        {
            WCHAR buffer[MAX_PATH];
            GetWindowTextW(hWnd, buffer, MAX_PATH);
            if (PathIsDirectoryW(buffer)) {
                PathCchCanonicalize(pData->currentPath, MAX_PATH, buffer);
                UpdateTabTitle(pData);
                SetWindowTextW(hWnd, pData->currentPath);
                ListDirectory(pData);
                UpdateSortMark(pData);
                SetFocus(pData->hList);
            }
            else {
                SHELLEXECUTEINFOW sei = { sizeof(sei) };
                sei.fMask = SEE_MASK_NOCLOSEPROCESS;
                sei.hwnd = GetParent(hWnd);
                sei.lpVerb = L"open";
                sei.lpFile = buffer;
                sei.lpDirectory = pData->currentPath;
                sei.nShow = SW_SHOWNORMAL;
                ShellExecuteExW(&sei);
                pData->bProgrammaticPathChange = true;
                SetWindowTextW(hWnd, pData->currentPath);
                pData->bProgrammaticPathChange = false;
                SendMessageW(hWnd, EM_SETSEL, (WPARAM)-1, 0);
            }
            return 0;
        }
        else if (wParam == VK_DOWN) {
            SetFocus(pData->hList);
            if (ListView_GetItemCount(pData->hList) > 0) {
                if (ListView_GetNextItem(pData->hList, -1, LVNI_SELECTED) == -1) {
                    ListView_SetItemState(pData->hList, 0, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
                    ListView_EnsureVisible(pData->hList, 0, FALSE);
                }
            }
            return 0;
        }
        break;
    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, EditSubclassProc, uIdSubclass);
        break;
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK TabSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    static bool bTracking = false;

    switch (uMsg)
    {
    case WM_MOUSEMOVE:
    {
        if (!bTracking) {
            TRACKMOUSEEVENT tme = { sizeof(tme) };
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hWnd;
            if (TrackMouseEvent(&tme)) {
                bTracking = true;
            }
        }

        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        TCHITTESTINFO hti = { pt, 0 };
        int tabItem = TabCtrl_HitTest(hWnd, &hti);
        int currentlyHoveredTab = -1;

        if (tabItem != -1) {
            RECT rcItem;
            TabCtrl_GetItemRect(hWnd, tabItem, &rcItem);
            RECT rcClose = rcItem;
            rcClose.left = rcClose.right - 22;
            if (PtInRect(&rcClose, pt)) {
                currentlyHoveredTab = tabItem;
            }
        }

        if (currentlyHoveredTab != g_hoveredCloseButtonTab) {
            int oldHoveredTab = g_hoveredCloseButtonTab;
            g_hoveredCloseButtonTab = currentlyHoveredTab;

            if (oldHoveredTab != -1) {
                RECT rcItem;
                TabCtrl_GetItemRect(hWnd, oldHoveredTab, &rcItem);
                InvalidateRect(hWnd, &rcItem, FALSE);
            }
            if (g_hoveredCloseButtonTab != -1) {
                RECT rcItem;
                TabCtrl_GetItemRect(hWnd, g_hoveredCloseButtonTab, &rcItem);
                InvalidateRect(hWnd, &rcItem, FALSE);
            }
        }
        break;
    }

    case WM_MOUSELEAVE:
    {
        if (g_hoveredCloseButtonTab != -1) {
            int oldHoveredTab = g_hoveredCloseButtonTab;
            g_hoveredCloseButtonTab = -1;
            RECT rcItem;
            TabCtrl_GetItemRect(hWnd, oldHoveredTab, &rcItem);
            InvalidateRect(hWnd, &rcItem, FALSE);
        }
        bTracking = false;
        return 0;
    }

    case WM_LBUTTONDOWN:
    {
        TCHITTESTINFO hti;
        hti.pt.x = GET_X_LPARAM(lParam);
        hti.pt.y = GET_Y_LPARAM(lParam);

        int iTab = TabCtrl_HitTest(hWnd, &hti);
        if (iTab != -1)
        {
            RECT rcItem;
            TabCtrl_GetItemRect(hWnd, iTab, &rcItem);
            RECT rcClose = rcItem;
            rcClose.left = rcClose.right - 22;

            if (PtInRect(&rcClose, hti.pt))
            {
                CloseTab(GetParent(hWnd), iTab);
                return 0;
            }
        }
        break;
    }
    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, TabSubclassProc, uIdSubclass);
        break;
    }

    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void PerformIncrementalSearch(ExplorerTabData* pData) {
    if (pData->searchString.empty()) return;
    int currentIndex = ListView_GetNextItem(pData->hList, -1, LVNI_FOCUSED);
    int startIndex = (currentIndex == -1) ? 0 : currentIndex;
    for (size_t i = startIndex + 1; i < pData->fileData.size(); ++i) {
        if (StrCmpNIW(pData->fileData[i].findData.cFileName, pData->searchString.c_str(), pData->searchString.length()) == 0) { // ★修正
            ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_SetItemState(pData->hList, (int)i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_FOCUSED | LVIS_SELECTED);
            ListView_EnsureVisible(pData->hList, (int)i, FALSE);
            return;
        }
    }
    for (int i = 0; i <= startIndex; ++i) {
        if (StrCmpNIW(pData->fileData[i].findData.cFileName, pData->searchString.c_str(), pData->searchString.length()) == 0) { // ★修正
            ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_SetItemState(pData->hList, i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_EnsureVisible(pData->hList, i, FALSE);
            return;
        }
    }
}

LRESULT CALLBACK ListSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    ExplorerTabData* pData = GetTabDataFromChild(hWnd);
    if (!pData) return DefSubclassProc(hWnd, uMsg, wParam, lParam);

    switch (uMsg) {
    case WM_CHAR:
    {
        if (pData->searchTimerId) {
            KillTimer(GetParent(hWnd), pData->searchTimerId);
        }
        if (wParam >= 32) {
            pData->searchString += static_cast<WCHAR>(wParam);
            PerformIncrementalSearch(pData);
        }
        pData->searchTimerId = SetTimer(GetParent(hWnd), SEARCH_TIMER_ID, 1000, nullptr);
        return 0;
    }
    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, ListSubclassProc, uIdSubclass);
        break;
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_NCCALCSIZE:
    {
        if (wParam == TRUE) {
            return 0;
        }
        break;
    }
    case WM_NCHITTEST:
    {
        POINT ptScreen = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        POINT ptClient = ptScreen;
        ScreenToClient(hWnd, &ptClient);

        if (PtInRect(&g_rcMinButton, ptClient)) {
            if (g_pressedButton != CaptionButtonState::None && g_pressedButton != CaptionButtonState::Min) {
                g_pressedButton = CaptionButtonState::None;
                InvalidateRect(hWnd, NULL, TRUE);
            }
            return HTMINBUTTON;
        }
        if (PtInRect(&g_rcMaxButton, ptClient)) {
            if (g_pressedButton != CaptionButtonState::None && g_pressedButton != CaptionButtonState::Max) {
                g_pressedButton = CaptionButtonState::None;
                InvalidateRect(hWnd, NULL, TRUE);
            }
            return HTMAXBUTTON;
        }
        if (PtInRect(&g_rcCloseButton, ptClient)) {
            if (g_pressedButton != CaptionButtonState::None && g_pressedButton != CaptionButtonState::Close) {
                g_pressedButton = CaptionButtonState::None;
                InvalidateRect(hWnd, NULL, TRUE);
            }
            return HTCLOSE;
        }

        LRESULT result;
        if (DwmDefWindowProc(hWnd, msg, wParam, lParam, &result)) {
            return result;
        }

        RECT rcWindow;
        GetWindowRect(hWnd, &rcWindow);

        if (!IsZoomed(hWnd) && !IsIconic(hWnd)) {
            int borderWidth = 8;
            if (ptScreen.x >= rcWindow.left && ptScreen.x < rcWindow.left + borderWidth && ptScreen.y >= rcWindow.top && ptScreen.y < rcWindow.top + borderWidth) return HTTOPLEFT;
            if (ptScreen.x < rcWindow.right && ptScreen.x >= rcWindow.right - borderWidth && ptScreen.y >= rcWindow.top && ptScreen.y < rcWindow.top + borderWidth) return HTTOPRIGHT;
            if (ptScreen.x >= rcWindow.left && ptScreen.x < rcWindow.left + borderWidth && ptScreen.y < rcWindow.bottom && ptScreen.y >= rcWindow.bottom - borderWidth) return HTBOTTOMLEFT;
            if (ptScreen.x < rcWindow.right && ptScreen.x >= rcWindow.right - borderWidth && ptScreen.y < rcWindow.bottom && ptScreen.y >= rcWindow.bottom - borderWidth) return HTBOTTOMRIGHT;
            if (ptScreen.x >= rcWindow.left && ptScreen.x < rcWindow.left + borderWidth) return HTLEFT;
            if (ptScreen.x < rcWindow.right && ptScreen.x >= rcWindow.right - borderWidth) return HTRIGHT;
            if (ptScreen.y >= rcWindow.top && ptScreen.y < rcWindow.top + borderWidth) return HTTOP;
            if (ptScreen.y < rcWindow.bottom && ptScreen.y >= rcWindow.bottom - borderWidth) return HTBOTTOM;
        }

        if (ptClient.y < 32) {
            TCHITTESTINFO hti = { ptClient, 0 };
            int tabItem = TabCtrl_HitTest(hTab, &hti);
            if (tabItem == -1) {
                RECT rcAddButton;
                GetWindowRect(hAddButton, &rcAddButton);
                MapWindowPoints(HWND_DESKTOP, hWnd, (LPPOINT)&rcAddButton, 2);
                if (!PtInRect(&rcAddButton, ptClient)) {
                    return HTCAPTION;
                }
            }
        }
        return HTCLIENT;
    }
    case WM_NCLBUTTONDOWN:
    {
        g_pressedButton = CaptionButtonState::None;
        switch (wParam) {
        case HTMINBUTTON:
            g_pressedButton = CaptionButtonState::Min;
            break;
        case HTMAXBUTTON:
            g_pressedButton = CaptionButtonState::Max;
            break;
        case HTCLOSE:
            g_pressedButton = CaptionButtonState::Close;
            break;
        }
        if (g_pressedButton != CaptionButtonState::None) {
            InvalidateRect(hWnd, NULL, TRUE);
            return 0;
        }
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    case WM_NCLBUTTONUP:
    {
        if (g_pressedButton != CaptionButtonState::None) {
            g_pressedButton = CaptionButtonState::None;
            InvalidateRect(hWnd, NULL, TRUE);
        }
        switch (wParam) {
        case HTMINBUTTON:
            SendMessage(hWnd, WM_SYSCOMMAND, SC_MINIMIZE, lParam);
            return 0;
        case HTMAXBUTTON:
            SendMessage(hWnd, WM_SYSCOMMAND, IsZoomed(hWnd) ? SC_RESTORE : SC_MAXIMIZE, lParam);
            return 0;
        case HTCLOSE:
            SendMessage(hWnd, WM_SYSCOMMAND, SC_CLOSE, lParam);
            return 0;
        }
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }

    case WM_CREATE:
    {
        hFont = CreateFontW(20, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, L"Segoe UI");
        MARGINS margins = { 0, 0, 32, 0 };
        DwmExtendFrameIntoClientArea(hWnd, &margins);
        SetWindowPos(hWnd, NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER);
        hTab = CreateWindowW(WC_TABCONTROLW, 0, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TCS_OWNERDRAWFIXED | TCS_BUTTONS,
            0, 0, 0, 0, hWnd, (HMENU)IDC_TAB, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
        SendMessage(hTab, WM_SETFONT, (WPARAM)hFont, TRUE);
        TabCtrl_SetPadding(hTab, 18, 4);
        SetWindowSubclass(hTab, TabSubclassProc, 0, 0);
        hAddButton = CreateWindowW(L"BUTTON", L"+", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            0, 0, 24, 24, hWnd, (HMENU)IDC_ADD_TAB_BUTTON, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
        SendMessage(hAddButton, WM_SETFONT, (WPARAM)hFont, TRUE);
        SHFILEINFO sfi = { 0 };
        hImgList = (HIMAGELIST)SHGetFileInfo(L"C:\\", 0, &sfi, sizeof(sfi), SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
        WCHAR initialPath[MAX_PATH];
        GetCurrentDirectoryW(MAX_PATH, initialPath);
        AddNewTab(hWnd, initialPath);
        UpdateCaptionButtonsRect(hWnd);
    }
    break;

    case WM_MOUSEMOVE:
    {
        if (!g_isTrackingMouse) {
            TRACKMOUSEEVENT tme = { sizeof(tme) };
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hWnd;
            if (TrackMouseEvent(&tme)) {
                g_isTrackingMouse = true;
            }
        }
        if (g_pressedButton != CaptionButtonState::None && GetAsyncKeyState(VK_LBUTTON) == 0) {
            g_pressedButton = CaptionButtonState::None;
            InvalidateRect(hWnd, NULL, TRUE);
        }
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        CaptionButtonState oldHovered = g_hoveredButton;
        g_hoveredButton = CaptionButtonState::None;
        if (PtInRect(&g_rcMinButton, pt)) g_hoveredButton = CaptionButtonState::Min;
        else if (PtInRect(&g_rcMaxButton, pt)) g_hoveredButton = CaptionButtonState::Max;
        else if (PtInRect(&g_rcCloseButton, pt)) g_hoveredButton = CaptionButtonState::Close;
        if (oldHovered != g_hoveredButton) {
            RECT redrawRect = g_rcMinButton;
            UnionRect(&redrawRect, &redrawRect, &g_rcMaxButton);
            UnionRect(&redrawRect, &redrawRect, &g_rcCloseButton);
            InvalidateRect(hWnd, &redrawRect, TRUE);
        }
    }
    break;

    case WM_MOUSELEAVE:
    {
        g_isTrackingMouse = false;
        if (g_hoveredButton != CaptionButtonState::None || g_pressedButton != CaptionButtonState::None) {
            g_hoveredButton = CaptionButtonState::None;
            g_pressedButton = CaptionButtonState::None;
            RECT redrawRect = g_rcMinButton;
            UnionRect(&redrawRect, &redrawRect, &g_rcMaxButton);
            UnionRect(&redrawRect, &redrawRect, &g_rcCloseButton);
            InvalidateRect(hWnd, &redrawRect, TRUE);
        }
    }
    break;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        DrawCaptionButtons(hdc, hWnd);
        EndPaint(hWnd, &ps);
        return 0;
    }

    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT lpdis = (LPDRAWITEMSTRUCT)lParam;
        if (lpdis->CtlID == IDC_ADD_TAB_BUTTON) {
            HTHEME hTheme = OpenThemeData(lpdis->hwndItem, L"BUTTON");
            if (hTheme) {
                HDC hdc = lpdis->hDC;
                RECT rc = lpdis->rcItem;
                int iState = PBS_NORMAL;
                if (lpdis->itemState & ODS_SELECTED) iState = PBS_PRESSED;
                else if (lpdis->itemState & ODS_HOTLIGHT) iState = PBS_HOT;
                DrawThemeBackground(hTheme, hdc, BP_PUSHBUTTON, iState, &rc, NULL);
                WCHAR szText[2] = L"+";
                DTTOPTS dttOpts = { sizeof(dttOpts) };
                dttOpts.dwFlags = DTT_COMPOSITED | DTT_TEXTCOLOR;
                dttOpts.crText = RGB(0, 0, 0);
                DrawThemeTextEx(hTheme, hdc, BP_PUSHBUTTON, iState, szText, -1, DT_CENTER | DT_VCENTER | DT_SINGLELINE, &rc, &dttOpts);
                CloseThemeData(hTheme);
            }
            return TRUE;
        }
        if (lpdis->CtlID == IDC_TAB) {
            HDC hdc = lpdis->hDC;
            RECT rcItem = lpdis->rcItem;
            int iItem = lpdis->itemID;
            int iSel = TabCtrl_GetCurSel(hTab);
            HTHEME hTheme = OpenThemeData(lpdis->hwndItem, L"TAB");
            if (hTheme)
            {
                int iState = (iItem == iSel) ? TIS_SELECTED : TIS_NORMAL;
                DrawThemeBackground(hTheme, hdc, TABP_TABITEM, iState, &rcItem, NULL);
                WCHAR szText[MAX_PATH];
                TCITEMW tci = { 0 };
                tci.mask = TCIF_TEXT;
                tci.pszText = szText;
                tci.cchTextMax = ARRAYSIZE(szText);
                TabCtrl_GetItem(hTab, iItem, &tci);
                RECT rcContent;
                GetThemeBackgroundContentRect(hTheme, hdc, TABP_TABITEM, iState, &rcItem, &rcContent);
                RECT rcClose = rcContent;
                rcClose.left = rcContent.right - 20;
                rcClose.right = rcContent.right;
                rcClose.top = rcContent.top + (rcContent.bottom - rcContent.top - 16) / 2;
                rcClose.bottom = rcClose.top + 16;
                RECT rcText = rcContent;
                rcText.right = rcClose.left - 4;
                rcText.left += 5;
                DTTOPTS dttOpts = { sizeof(dttOpts) };
                dttOpts.dwFlags = DTT_COMPOSITED | DTT_TEXTCOLOR;
                DrawThemeTextEx(hTheme, hdc, TABP_TABITEM, iState, szText, -1, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS, &rcText, &dttOpts);
                int iCloseState = TCBS_NORMAL;
                if (iItem == g_hoveredCloseButtonTab) iCloseState = TCBS_HOT;
                DrawThemeBackground(hTheme, hdc, TABP_CLOSEBUTTON, iCloseState, &rcClose, NULL);
                CloseThemeData(hTheme);
            }
            if (lpdis->itemState & ODS_FOCUS) {
                DrawFocusRect(hdc, &rcItem);
            }
            return TRUE;
        }
    }
    break;

    case WM_TIMER:
    {
        if (wParam == SEARCH_TIMER_ID) {
            ExplorerTabData* pData = GetCurrentTabData();
            if (pData) {
                KillTimer(hWnd, pData->searchTimerId);
                pData->searchTimerId = 0;
                pData->searchString.clear();
            }
        }
    }
    break;

    case WM_ACTIVATE:
    {
        if (LOWORD(wParam) != WA_INACTIVE) {
            ExplorerTabData* pData = GetCurrentTabData();
            if (pData) SetFocus(pData->hList);
        }
        InvalidateRect(hWnd, NULL, TRUE);
    }
    break;

    case WM_SIZE:
        UpdateCaptionButtonsRect(hWnd);
        UpdateLayout(hWnd);
        break;

    case WM_COMMAND:
    {
        ExplorerTabData* pData = GetCurrentTabData();
        if (!pData && LOWORD(wParam) != IDC_ADD_TAB_BUTTON) break;
        if (HIWORD(wParam) == EN_UPDATE && GetDlgCtrlID((HWND)lParam) >= EDIT_ID_BASE) {
            if (!pData->bProgrammaticPathChange) {
                WCHAR tempPath[MAX_PATH];
                GetWindowTextW(pData->hEdit, tempPath, MAX_PATH);
                if (PathIsDirectoryW(tempPath)) {
                    PathCchCanonicalize(pData->currentPath, MAX_PATH, tempPath);
                    ListDirectory(pData);
                }
            }
        }
        switch (LOWORD(wParam)) {
        case IDC_ADD_TAB_BUTTON:
        {
            ExplorerTabData* currentData = GetCurrentTabData();
            if (currentData) {
                AddNewTab(hWnd, currentData->currentPath);
            }
            else {
                WCHAR initialPath[MAX_PATH];
                GetCurrentDirectoryW(MAX_PATH, initialPath);
                AddNewTab(hWnd, initialPath);
            }
            break;
        }
        case ID_ACCELERATOR_UP:
            NavigateUpAndSelect(hWnd);
            break;
        case ID_ACCELERATOR_ADDRESS:
            SetFocus(pData->hEdit);
            SendMessage(pData->hEdit, EM_SETSEL, 0, -1);
            break;
        case ID_ACCELERATOR_TAB:
        {
            HWND hFocus = GetFocus();
            if (hFocus == pData->hEdit) SetFocus(pData->hList);
            else {
                SetFocus(pData->hEdit);
                SendMessage(pData->hEdit, EM_SETSEL, 0, -1);
            }
        }
        break;
        case ID_ACCELERATOR_REFRESH:
            ListDirectory(pData);
            UpdateSortMark(pData);
            break;
        case ID_ACCELERATOR_NEW_TAB:
            AddNewTab(hWnd, pData->currentPath);
            break;
        case ID_ACCELERATOR_CLOSE_TAB:
            CloseTab(hWnd, TabCtrl_GetCurSel(hTab));
            break;
        case ID_ACCELERATOR_NEXT_TAB:
        {
            int tabCount = TabCtrl_GetItemCount(hTab);
            if (tabCount > 1) {
                int currentIndex = TabCtrl_GetCurSel(hTab);
                int nextIndex = (currentIndex + 1) % tabCount;
                SwitchTab(hWnd, nextIndex);
            }
            break;
        }
        case ID_ACCELERATOR_PREV_TAB:
        {
            int tabCount = TabCtrl_GetItemCount(hTab);
            if (tabCount > 1) {
                int currentIndex = TabCtrl_GetCurSel(hTab);
                int prevIndex = (currentIndex - 1 + tabCount) % tabCount;
                SwitchTab(hWnd, prevIndex);
            }
            break;
        }
        }
    }
    break;

    case WM_NOTIFY:
    {
        LPNMHDR lpnmh = (LPNMHDR)lParam;
        if (lpnmh->idFrom == IDC_TAB) {
            if (lpnmh->code == TCN_SELCHANGE) {
                SwitchTab(hWnd, TabCtrl_GetCurSel(hTab));
            }
        }
        else {
            ExplorerTabData* pData = GetTabDataFromChild(lpnmh->hwndFrom);
            if (pData) {
                switch (lpnmh->code) {
                case LVN_GETDISPINFO:
                {
                    NMLVDISPINFO* plvdi = (NMLVDISPINFO*)lParam;
                    LVITEM* pItem = &(plvdi->item);
                    if (pItem->iItem >= (int)pData->fileData.size()) break;
                    const FileInfo& info = pData->fileData[pItem->iItem]; // ★修正
                    const WIN32_FIND_DATAW& fd = info.findData; // ★修正
                    if (pItem->mask & LVIF_TEXT) {
                        switch (pItem->iSubItem) {
                        case 0: StringCchCopy(pItem->pszText, pItem->cchTextMax, fd.cFileName); break;
                        case 1:
                            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                                StringCchCopy(pItem->pszText, pItem->cchTextMax, L"");
                            }
                            else {
                                ULARGE_INTEGER sz = { fd.nFileSizeLow, fd.nFileSizeHigh };
                                StrFormatByteSizeW(sz.QuadPart, pItem->pszText, pItem->cchTextMax);
                            }
                            break;
                        case 2:
                        case 3:
                            if (wcscmp(fd.cFileName, L"..") != 0) {
                                const FILETIME& ft = (pItem->iSubItem == 2) ? fd.ftLastWriteTime : fd.ftCreationTime;
                                FileTimeToString(ft, pItem->pszText, pItem->cchTextMax);
                            }
                            else {
                                pItem->pszText[0] = L'\0';
                            }
                            break;
                        case 4: // ★追加: 所有者情報の表示
                            StringCchCopy(pItem->pszText, pItem->cchTextMax, info.owner.c_str());
                            break;
                        }
                    }
                    if (pItem->mask & LVIF_IMAGE) {
                        WCHAR fullpath[MAX_PATH];
                        PathCombineW(fullpath, pData->currentPath, fd.cFileName);
                        SHFILEINFO sfi = { 0 };
                        SHGetFileInfo(fullpath, fd.dwFileAttributes, &sfi, sizeof(sfi), SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
                        pItem->iImage = sfi.iIcon;
                    }
                }
                break;
                case LVN_COLUMNCLICK:
                {
                    LPNMLISTVIEW pnmv = (LPNMLISTVIEW)lParam;
                    int clickedColumn = pnmv->iSubItem;
                    if (clickedColumn == pData->sortColumn) {
                        pData->sortAscending = !pData->sortAscending;
                    }
                    else {
                        pData->sortColumn = clickedColumn;
                        pData->sortAscending = true;
                    }
                    std::sort(pData->fileData.begin(), pData->fileData.end(), [&](const FileInfo& a, const FileInfo& b) { // ★修正
                        return CompareFunction(a, b, pData);
                        });
                    UpdateSortMark(pData);
                    InvalidateRect(pData->hList, NULL, TRUE);
                }
                break;
                case LVN_KEYDOWN:
                {
                    LPNMLVKEYDOWN lpnmkv = (LPNMLVKEYDOWN)lParam;
                    switch (lpnmkv->wVKey) {
                    case VK_RETURN:
                        NavigateToSelected(pData);
                        break;
                    case VK_BACK:
                        NavigateUpAndSelect(hWnd);
                        break;
                    case VK_DELETE:
                        DoDelete(hWnd, GetKeyState(VK_SHIFT) & 0x8000);
                        break;
                    case VK_APPS:
                    {
                        POINT pt = {};
                        int sel = ListView_GetNextItem(pData->hList, -1, LVNI_SELECTED);
                        if (sel != -1) {
                            RECT rc;
                            ListView_GetItemRect(pData->hList, sel, &rc, LVIR_BOUNDS);
                            pt.x = rc.left;
                            pt.y = rc.bottom;
                            ClientToScreen(pData->hList, &pt);
                            ShowContextMenu(hWnd, pData, sel, pt);
                        }
                    }
                    break;
                    case 'A':
                        if (GetKeyState(VK_CONTROL) & 0x8000) {
							int itemCount = ListView_GetItemCount(pData->hList);
							for (int i = 0; i < itemCount; ++i) {
								ListView_SetItemState(pData->hList, i, LVIS_SELECTED, LVIS_SELECTED);
							}
                        }
                        break;
                    case 'C':
                        if (GetKeyState(VK_CONTROL) & 0x8000) {
                            DoCopyCut(hWnd, false);
                        }
                        break;
                    case 'X':
                        if (GetKeyState(VK_CONTROL) & 0x8000) {
                            DoCopyCut(hWnd, true);
                        }
                        break;
                    case 'V':
                        if (GetKeyState(VK_CONTROL) & 0x8000) {
                            DoPaste(hWnd);
                        }
                        break;
                    }
                }
                break;
                case NM_DBLCLK:
                    NavigateToSelected(pData);
                    break;
                case NM_RCLICK:
                {
                    LPNMITEMACTIVATE lpnmitem = (LPNMITEMACTIVATE)lParam;
                    POINT pt = lpnmitem->ptAction;
                    ClientToScreen(pData->hList, &pt);
                    ShowContextMenu(hWnd, pData, lpnmitem->iItem, pt);
                }
                break;
                }
            }
        }
    }
    break;

    case WM_DESTROY:
        for (const auto& tab : g_tabs) {
            if (tab->searchTimerId) KillTimer(hWnd, tab->searchTimerId);
        }
        DeleteObject(hFont);
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

void AddNewTab(HWND hWnd, LPCWSTR initialPath) {
    auto pData = std::make_unique<ExplorerTabData>();
    if (!pData) return;

    pData->sortColumn = 0;
    pData->sortAscending = true;
    pData->searchTimerId = 0;
    pData->bProgrammaticPathChange = false;
    StringCchCopy(pData->currentPath, MAX_PATH, initialPath);

    int tabIndex = (int)g_tabs.size();

    pData->hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | ES_AUTOHSCROLL,
        0, 0, 0, 0,
        hWnd, (HMENU)(EDIT_ID_BASE + tabIndex), GetModuleHandle(NULL), NULL);

    pData->hList = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | LVS_REPORT | LVS_OWNERDATA,
        0, 0, 0, 0,
        hWnd, (HMENU)(LIST_ID_BASE + tabIndex), GetModuleHandle(NULL), NULL);

    if (!pData->hEdit || !pData->hList) return;

    SetWindowTheme(pData->hList, L"Explorer", NULL);

    SetWindowSubclass(pData->hEdit, EditSubclassProc, 0, (DWORD_PTR)pData.get());
    SetWindowSubclass(pData->hList, ListSubclassProc, 0, (DWORD_PTR)pData.get());

    SendMessage(pData->hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(pData->hList, WM_SETFONT, (WPARAM)hFont, TRUE);
    SHAutoComplete(pData->hEdit, SHACF_FILESYS_DIRS | SHACF_AUTOSUGGEST_FORCE_ON | SHACF_AUTOAPPEND_FORCE_ON);
    ListView_SetExtendedListViewStyle(pData->hList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);
    ListView_SetImageList(pData->hList, hImgList, LVSIL_SMALL);

    LVCOLUMNW lvc = { LVCF_TEXT | LVCF_WIDTH | LVCF_FMT };
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 250; lvc.pszText = (LPWSTR)L"名前"; ListView_InsertColumn(pData->hList, 0, &lvc);
    lvc.fmt = LVCFMT_RIGHT; lvc.cx = 100; lvc.pszText = (LPWSTR)L"サイズ"; ListView_InsertColumn(pData->hList, 1, &lvc);
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 150; lvc.pszText = (LPWSTR)L"更新日時"; ListView_InsertColumn(pData->hList, 2, &lvc);
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 150; lvc.pszText = (LPWSTR)L"作成日時"; ListView_InsertColumn(pData->hList, 3, &lvc);
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 150; lvc.pszText = (LPWSTR)L"所有者"; ListView_InsertColumn(pData->hList, 4, &lvc); // ★追加

    SetWindowTextW(pData->hEdit, pData->currentPath);
    ListDirectory(pData.get());
    UpdateSortMark(pData.get());

    TCITEMW tci = { TCIF_TEXT };
    WCHAR tabText[MAX_PATH];
    StringCchCopyW(tabText, MAX_PATH, PathFindFileNameW(pData->currentPath));
    if (wcslen(tabText) == 0) {
        StringCchCopyW(tabText, MAX_PATH, pData->currentPath);
        PathRemoveBackslashW(tabText);
    }
    tci.pszText = tabText;
    TabCtrl_InsertItem(hTab, tabIndex, &tci);

    g_tabs.push_back(std::move(pData));
    SwitchTab(hWnd, tabIndex);
}

void CloseTab(HWND hWnd, int index) {
    if (index < 0 || index >= (int)g_tabs.size()) return;
    if (g_tabs.size() <= 1) {
        DestroyWindow(hWnd);
        return;
    }
    ExplorerTabData* pData = g_tabs[index].get();
    DestroyWindow(pData->hEdit);
    DestroyWindow(pData->hList);

    g_tabs.erase(g_tabs.begin() + index);
    TabCtrl_DeleteItem(hTab, index);

    int newSel = (index >= TabCtrl_GetItemCount(hTab)) ? (TabCtrl_GetItemCount(hTab) - 1) : index;
    if (newSel < 0) newSel = 0;

    SwitchTab(hWnd, newSel);
}

void SwitchTab(HWND hWnd, int newIndex) {
    TabCtrl_SetCurSel(hTab, newIndex);
    for (size_t i = 0; i < g_tabs.size(); ++i) {
        BOOL show = ((int)i == newIndex);
        ShowWindow(g_tabs[i]->hEdit, show ? SW_SHOW : SW_HIDE);
        ShowWindow(g_tabs[i]->hList, show ? SW_SHOW : SW_HIDE);
    }
    if (newIndex >= 0 && newIndex < (int)g_tabs.size()) {
        SetFocus(g_tabs[newIndex]->hList);
    }
    UpdateLayout(hWnd);
}

void UpdateLayout(HWND hWnd) {
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);

    constexpr int BUTTON_WIDTH = 22;
    constexpr int BUTTON_MARGIN = 4;
    constexpr int TITLEBAR_HEIGHT = 32;

    int tabHeight = 30;
    int tabY = (TITLEBAR_HEIGHT - tabHeight) / 2;
    int buttonY = (TITLEBAR_HEIGHT - BUTTON_WIDTH) / 2;

    int tabCount = TabCtrl_GetItemCount(hTab);
    int buttonX = 0;

    if (tabCount > 0) {
        HWND hUpDown = FindWindowEx(hTab, NULL, UPDOWN_CLASS, NULL);
        bool scrollButtonsVisible = (hUpDown != NULL && IsWindowVisible(hUpDown));
        RECT rcLastTab;
        TabCtrl_GetItemRect(hTab, tabCount - 1, &rcLastTab);

        constexpr int CAPTION_BUTTONS_WIDTH = 47 * 3;
        int availableWidth = rcClient.right - CAPTION_BUTTONS_WIDTH;

        int requiredWidth = rcLastTab.right + BUTTON_MARGIN + BUTTON_WIDTH + BUTTON_MARGIN;
        if (scrollButtonsVisible || requiredWidth >= availableWidth) {
            constexpr int RESERVED_AREA_WIDTH = BUTTON_WIDTH + BUTTON_MARGIN * 2;
            int tabWidth = availableWidth - RESERVED_AREA_WIDTH;
            if (tabWidth < 0) tabWidth = 0;

            MoveWindow(hTab, 0, tabY, tabWidth, tabHeight, TRUE);
            buttonX = tabWidth + BUTTON_MARGIN;
        }
        else {
            buttonX = rcLastTab.right + BUTTON_MARGIN;
            MoveWindow(hTab, 0, tabY, availableWidth, tabHeight, TRUE);
        }
    }
    else {
        buttonX = BUTTON_MARGIN;
        MoveWindow(hTab, 0, tabY, rcClient.right, tabHeight, TRUE);
    }

    MoveWindow(hAddButton, buttonX, buttonY, BUTTON_WIDTH, BUTTON_WIDTH, TRUE);
    SetWindowPos(hAddButton, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

    int curSel = TabCtrl_GetCurSel(hTab);
    if (curSel != -1) {
        ExplorerTabData* pData = GetCurrentTabData();
        if (pData) {
            int contentTop = TITLEBAR_HEIGHT;
            MoveWindow(pData->hEdit, 0, contentTop, rcClient.right, 24, TRUE);
            MoveWindow(pData->hList, 0, contentTop + 24, rcClient.right, rcClient.bottom - contentTop - 24, TRUE);
            UpdateWindow(pData->hList);
        }
    }
    InvalidateRect(hWnd, NULL, TRUE);
}

void UpdateCaptionButtonsRect(HWND hWnd) {
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);

    constexpr int BTN_WIDTH = 47;
    constexpr int BTN_HEIGHT = 32;

    g_rcCloseButton.top = 0;
    g_rcCloseButton.bottom = BTN_HEIGHT;
    g_rcCloseButton.right = rcClient.right;
    g_rcCloseButton.left = rcClient.right - BTN_WIDTH;

    g_rcMaxButton.top = 0;
    g_rcMaxButton.bottom = BTN_HEIGHT;
    g_rcMaxButton.right = g_rcCloseButton.left;
    g_rcMaxButton.left = g_rcCloseButton.left - BTN_WIDTH;

    g_rcMinButton.top = 0;
    g_rcMinButton.bottom = BTN_HEIGHT;
    g_rcMinButton.right = g_rcMaxButton.left;
    g_rcMinButton.left = g_rcMaxButton.left - BTN_WIDTH;
}

void DrawCaptionButtons(HDC hdc, HWND hWnd) {
    HTHEME hTheme = OpenThemeData(hWnd, L"WINDOW");
    if (hTheme) {
        int iState;
        BOOL bActive = (GetActiveWindow() == hWnd);

        iState = MINBS_NORMAL;
        if (!bActive) iState = MINBS_INACTIVE;
        else if (g_pressedButton == CaptionButtonState::Min) iState = MINBS_PUSHED;
        else if (g_hoveredButton == CaptionButtonState::Min) iState = MINBS_HOT;
        DrawThemeBackground(hTheme, hdc, WP_MINBUTTON, iState, &g_rcMinButton, NULL);

        if (IsZoomed(hWnd)) {
            iState = RBS_NORMAL;
            if (!bActive) iState = RBS_INACTIVE;
            else if (g_pressedButton == CaptionButtonState::Max) iState = RBS_PUSHED;
            else if (g_hoveredButton == CaptionButtonState::Max) iState = RBS_HOT;
            DrawThemeBackground(hTheme, hdc, WP_RESTOREBUTTON, iState, &g_rcMaxButton, NULL);
        }
        else {
            iState = MAXBS_NORMAL;
            if (!bActive) iState = MAXBS_INACTIVE;
            else if (g_pressedButton == CaptionButtonState::Max) iState = MAXBS_PUSHED;
            else if (g_hoveredButton == CaptionButtonState::Max) iState = MAXBS_HOT;
            DrawThemeBackground(hTheme, hdc, WP_MAXBUTTON, iState, &g_rcMaxButton, NULL);
        }

        iState = CBS_NORMAL;
        if (!bActive) iState = CBS_INACTIVE;
        else if (g_pressedButton == CaptionButtonState::Close) iState = CBS_PUSHED;
        else if (g_hoveredButton == CaptionButtonState::Close) iState = CBS_HOT;
        DrawThemeBackground(hTheme, hdc, WP_CLOSEBUTTON, iState, &g_rcCloseButton, NULL);

        CloseThemeData(hTheme);
    }
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPWSTR pCmdLine, int nCmdShow) {
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    INITCOMMONCONTROLSEX ic = { sizeof(ic), ICC_LISTVIEW_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&ic);
    MSG msg;
    WNDCLASSW wndclass = {
        CS_HREDRAW | CS_VREDRAW,
        WndProc,
        0,0,
        hInstance,
        LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)),
        LoadCursor(0,IDC_ARROW),
        (HBRUSH)GetStockObject(BLACK_BRUSH),
        0,
        szClassName
    };
    RegisterClassW(&wndclass);
    HWND hWnd = CreateWindowEx(0,
        szClassName, L"タブファイラ",
        WS_POPUP | WS_THICKFRAME | WS_SYSMENU | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_CLIPCHILDREN,
        CW_USEDEFAULT, 0, 900, 600,
        0, 0, hInstance, 0);
    ShowWindow(hWnd, SW_SHOWDEFAULT);
    UpdateWindow(hWnd);
    HACCEL hAccel = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1));
    while (GetMessage(&msg, 0, 0, 0)) {
        if (!TranslateAccelerator(hWnd, hAccel, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    CoUninitialize();
    return (int)msg.wParam;
}