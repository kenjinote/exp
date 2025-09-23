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

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "pathcch.lib")
#pragma comment(lib, "ole32.lib")

#define IDC_TAB 1000
#define EDIT_ID_BASE 1001
#define LIST_ID_BASE 2001
#define IDC_ADD_TAB_BUTTON 3001
#define SEARCH_TIMER_ID 1

// 各タブの状態を保持する構造体
struct ExplorerTabData {
    HWND hEdit;
    HWND hList;
    WCHAR currentPath[MAX_PATH];
    std::vector<WIN32_FIND_DATAW> fileData;
    int sortColumn;
    bool sortAscending;
    std::wstring searchString;
    UINT_PTR searchTimerId;
    bool bProgrammaticPathChange;
};

// --- グローバル変数 ---
WCHAR szClassName[] = L"exp";
static HWND hTab;
static HWND hAddButton;
static HFONT hFont;
static HFONT hMarlettFont;
static HIMAGELIST hImgList;
static std::vector<std::unique_ptr<ExplorerTabData>> g_tabs;
static int g_hoveredCloseButtonTab = -1; // <<< この行を追加

// --- 関数の前方宣言 ---
void SwitchTab(HWND hWnd, int newIndex);
void UpdateLayout(HWND hWnd);
void AddNewTab(HWND hWnd, LPCWSTR initialPath);
void CloseTab(HWND hWnd, int index);
LRESULT CALLBACK EditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK ListSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
void ListDirectory(ExplorerTabData* pData);
void UpdateSortMark(ExplorerTabData* pData);
void UpdateTabTitle(ExplorerTabData* pData); // <<< この行を追加

// --- 補助関数 (変更なし) ---
ExplorerTabData* GetCurrentTabData() {
    int sel = TabCtrl_GetCurSel(hTab);
    if (sel >= 0 && sel < g_tabs.size()) {
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

// --- ファイル操作関数 (変更なし) ---

bool CompareFunction(const WIN32_FIND_DATAW& a, const WIN32_FIND_DATAW& b, ExplorerTabData* pData) {
    if (wcscmp(a.cFileName, L"..") == 0) return true;
    if (wcscmp(b.cFileName, L"..") == 0) return false;
    const bool isDirA = (a.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
    const bool isDirB = (b.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY);
    if (isDirA != isDirB) {
        return isDirA > isDirB;
    }
    int result = 0;
    switch (pData->sortColumn) {
    case 0:
        result = StrCmpLogicalW(a.cFileName, b.cFileName);
        break;
    case 1:
    {
        ULARGE_INTEGER sizeA = { a.nFileSizeLow, a.nFileSizeHigh };
        ULARGE_INTEGER sizeB = { b.nFileSizeLow, b.nFileSizeHigh };
        if (sizeA.QuadPart < sizeB.QuadPart) result = -1;
        else if (sizeA.QuadPart > sizeB.QuadPart) result = 1;
        else result = 0;
        break;
    }
    case 2:
        result = CompareFileTime(&a.ftLastWriteTime, &b.ftLastWriteTime);
        break;
    case 3:
        result = CompareFileTime(&a.ftCreationTime, &b.ftCreationTime);
        break;
    }
    return pData->sortAscending ? (result < 0) : (result > 0);
}

void ListDirectory(ExplorerTabData* pData) {
    pData->fileData.clear();
    if (PathIsRootW(pData->currentPath) == FALSE) {
        WIN32_FIND_DATAW fd_up = { 0 };
        fd_up.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        StringCchCopyW(fd_up.cFileName, MAX_PATH, L"..");
        pData->fileData.push_back(fd_up);
    }
    WIN32_FIND_DATAW fd;
    WCHAR searchPath[MAX_PATH];
    PathCombineW(searchPath, pData->currentPath, L"*");
    HANDLE hFind = FindFirstFileW(searchPath, &fd);
    if (hFind != INVALID_HANDLE_VALUE) {
        do {
            if (wcscmp(fd.cFileName, L".") != 0 && wcscmp(fd.cFileName, L"..") != 0) {
                pData->fileData.push_back(fd);
            }
        } while (FindNextFileW(hFind, &fd));
        FindClose(hFind);
    }
    std::sort(pData->fileData.begin(), pData->fileData.end(), [&](const WIN32_FIND_DATAW& a, const WIN32_FIND_DATAW& b) {
        return CompareFunction(a, b, pData);
        });
    ListView_SetItemCountEx(pData->hList, pData->fileData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(pData->hList, NULL, TRUE);
}

void GetSelectedFilePathsW(ExplorerTabData* pData, std::vector<std::wstring>& paths) {
    paths.clear();
    int iItem = -1;
    while ((iItem = ListView_GetNextItem(pData->hList, iItem, LVNI_SELECTED)) != -1) {
        if (iItem < pData->fileData.size()) {
            const WIN32_FIND_DATAW& fd = pData->fileData[iItem];
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
            if (currentSel < pData->fileData.size()) {
                const WIN32_FIND_DATAW& fd = pData->fileData[currentSel];
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
    ExplorerTabData* pData = GetCurrentTabData(); // pDataを取得
    if (!pData) return;

    WCHAR previousFolderName[MAX_PATH];
    if (PathIsRootW(pData->currentPath) == FALSE) {
        StringCchCopyW(previousFolderName, MAX_PATH, PathFindFileNameW(pData->currentPath));
        PathCchRemoveFileSpec(pData->currentPath, MAX_PATH);

        UpdateTabTitle(pData); // ▼▼▼ 追加 ▼▼▼ タブ名を更新

        SetWindowTextW(pData->hEdit, pData->currentPath);
        ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED);
        ListDirectory(pData);
        UpdateSortMark(pData);
        for (size_t i = 0; i < pData->fileData.size(); ++i) {
            if (wcscmp(pData->fileData[i].cFileName, previousFolderName) == 0) {
                ListView_SetItemState(pData->hList, i, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
                ListView_EnsureVisible(pData->hList, i, FALSE);
                break;
            }
        }
    }
}

void NavigateToSelected(ExplorerTabData* pData) {
    int sel = ListView_GetNextItem(pData->hList, -1, LVNI_SELECTED);
    if (sel < 0 || sel >= pData->fileData.size()) {
        return;
    }
    const WIN32_FIND_DATAW& fd = pData->fileData[sel];
    if (wcscmp(fd.cFileName, L"..") == 0) {
        NavigateUpAndSelect(GetParent(pData->hList));
        return;
    }
    if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
        WCHAR newPath[MAX_PATH];
        PathCombineW(newPath, pData->currentPath, fd.cFileName);
        StringCchCopyW(pData->currentPath, MAX_PATH, newPath);

        UpdateTabTitle(pData); // ▼▼▼ 追加 ▼▼▼ タブ名を更新

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
    // pDataからタブのインデックスを検索
    int tabIndex = -1;
    for (size_t i = 0; i < g_tabs.size(); ++i) {
        if (g_tabs[i].get() == pData) {
            tabIndex = i;
            break;
        }
    }
    if (tabIndex == -1) return;

    // 新しいタブのテキスト（フォルダ名）を準備
    TCITEMW tci = { TCIF_TEXT };
    WCHAR tabText[MAX_PATH];
    StringCchCopyW(tabText, MAX_PATH, PathFindFileNameW(pData->currentPath));
    if (wcslen(tabText) == 0) { // ルートディレクトリの場合 (例: "C:\")
        StringCchCopyW(tabText, MAX_PATH, pData->currentPath);
        PathRemoveBackslashW(tabText);
    }
    tci.pszText = tabText;

    // タブのテキストを設定
    TabCtrl_SetItem(hTab, tabIndex, &tci);

    // ▼▼▼ 追加 ▼▼▼
    // タブ幅の変更後にレイアウトを更新し、「+」ボタンの位置を再計算する
    UpdateLayout(GetParent(hTab));
}

// ▼▼▼ この関数を新しく追加 ▼▼▼
void RebuildTabsFromData(int newSelection) {
    // UI上のタブを一旦すべて削除
    TabCtrl_DeleteAllItems(hTab);

    // g_tabs ベクターの現在の順序に基づいてタブを再作成
    for (size_t i = 0; i < g_tabs.size(); ++i) {
        TCITEMW tci = { TCIF_TEXT };
        WCHAR tabText[MAX_PATH];
        // パスからフォルダ名を取得してタブのテキストとする
        StringCchCopyW(tabText, MAX_PATH, PathFindFileNameW(g_tabs[i]->currentPath));
        if (wcslen(tabText) == 0) { // ルートディレクトリの場合
            StringCchCopyW(tabText, MAX_PATH, g_tabs[i]->currentPath);
            PathRemoveBackslashW(tabText);
        }
        tci.pszText = tabText;
        TabCtrl_InsertItem(hTab, i, &tci);
    }

    // 指定されたタブを選択状態にする
    if (newSelection >= 0 && newSelection < g_tabs.size()) {
        SwitchTab(GetParent(hTab), newSelection);
    }
}

// --- サブクラスプロシージャ (変更なし) ---

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

                UpdateTabTitle(pData); // ▼▼▼ 追加 ▼▼▼ タブ名を更新

                SetWindowTextW(hWnd, pData->currentPath);
                ListDirectory(pData);
                UpdateSortMark(pData);
                SetFocus(pData->hList);
            } else {
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
        int currentlyHoveredTab = -1; // 今回のイベントでホバーされているタブ

        if (tabItem != -1) {
            RECT rcItem;
            TabCtrl_GetItemRect(hWnd, tabItem, &rcItem);
            RECT rcClose = rcItem;
            rcClose.left = rcClose.right - 22;
            if (PtInRect(&rcClose, pt)) {
                currentlyHoveredTab = tabItem;
            }
        }

        // ▼▼▼ 変更 ▼▼▼ ホバー状態が変化したかチェック
        if (currentlyHoveredTab != g_hoveredCloseButtonTab) {
            int oldHoveredTab = g_hoveredCloseButtonTab;
            g_hoveredCloseButtonTab = currentlyHoveredTab; // 新しい状態を保存

            // 前にホバーしていたタブを再描画
            if (oldHoveredTab != -1) {
                RECT rcItem;
                TabCtrl_GetItemRect(hWnd, oldHoveredTab, &rcItem);
                InvalidateRect(hWnd, &rcItem, FALSE);
            }
            // 新しくホバーしたタブを再描画
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
        // ▼▼▼ 変更 ▼▼▼ マウスが離れたら状態をリセットして再描画
        if (g_hoveredCloseButtonTab != -1) {
            int oldHoveredTab = g_hoveredCloseButtonTab;
            g_hoveredCloseButtonTab = -1; // 状態をリセット
            RECT rcItem;
            TabCtrl_GetItemRect(hWnd, oldHoveredTab, &rcItem);
            InvalidateRect(hWnd, &rcItem, FALSE);
        }
        bTracking = false;
        return 0;
    }

    case WM_LBUTTONDOWN:
    {
        // この部分は変更なし
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
        if (StrCmpNIW(pData->fileData[i].cFileName, pData->searchString.c_str(), pData->searchString.length()) == 0) {
            ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_SetItemState(pData->hList, i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_EnsureVisible(pData->hList, i, FALSE);
            return;
        }
    }
    for (int i = 0; i <= startIndex; ++i) {
        if (StrCmpNIW(pData->fileData[i].cFileName, pData->searchString.c_str(), pData->searchString.length()) == 0) {
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


// --- メインウィンドウプロシージャ ---

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
    {
        hFont = CreateFontW(20, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, L"Segoe UI");

        hTab = CreateWindowW(WC_TABCONTROLW, 0, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TCS_OWNERDRAWFIXED,
            0, 0, 0, 0, hWnd, (HMENU)IDC_TAB, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
        SendMessage(hTab, WM_SETFONT, (WPARAM)hFont, TRUE);
        TabCtrl_SetPadding(hTab, 18, 4);
        SetWindowSubclass(hTab, TabSubclassProc, 0, 0);

        LOGFONTW lf = { 0 };
        lf.lfHeight = 16;
        lf.lfWeight = FW_NORMAL;
        lf.lfCharSet = SYMBOL_CHARSET;
        StringCchCopyW(lf.lfFaceName, LF_FACESIZE, L"Marlett");
        hMarlettFont = CreateFontIndirectW(&lf);

        hAddButton = CreateWindowW(L"BUTTON", L"+", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            0, 0, 24, 24, hWnd, (HMENU)IDC_ADD_TAB_BUTTON, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
        SendMessage(hAddButton, WM_SETFONT, (WPARAM)hFont, TRUE);

        SHFILEINFO sfi = { 0 };
        hImgList = (HIMAGELIST)SHGetFileInfo(L"C:\\", 0, &sfi, sizeof(sfi), SHGFI_SYSICONINDEX | SHGFI_SMALLICON);

        WCHAR initialPath[MAX_PATH];
        GetCurrentDirectoryW(MAX_PATH, initialPath);
        AddNewTab(hWnd, initialPath);
    }
    break;

    // ▼▼▼ 追加 ▼▼▼ タブバーの下に境界線を描画
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);

        RECT rcTab;
        GetWindowRect(hTab, &rcTab);
        MapWindowPoints(HWND_DESKTOP, hWnd, (LPPOINT)&rcTab, 2);

        // タブコントロールのすぐ下に線を描画
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(204, 204, 204));
        HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, rcTab.left, rcTab.bottom - 1, NULL);
        LineTo(hdc, rcTab.right, rcTab.bottom - 1);
        SelectObject(hdc, hOldPen);
        DeleteObject(hPen);

        EndPaint(hWnd, &ps);
    }
    break;

    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT lpdis = (LPDRAWITEMSTRUCT)lParam;
        if (lpdis->CtlID == IDC_TAB) {
            HDC hdc = lpdis->hDC;
            RECT rc = lpdis->rcItem;
            int iItem = lpdis->itemID;
            int iSel = TabCtrl_GetCurSel(hTab);

            COLORREF bgColor = (iItem == iSel) ? RGB(255, 255, 255) : RGB(240, 240, 240);
            COLORREF textColor = GetSysColor(COLOR_WINDOWTEXT);

            HBRUSH hBrush = CreateSolidBrush(bgColor);
            FillRect(hdc, &rc, hBrush);
            DeleteObject(hBrush);

            // 選択中のタブ以外は、右側に区切り線を描画
            if (iItem != iSel && iItem != TabCtrl_GetItemCount(hTab) - 1) {
                HPEN hPen = CreatePen(PS_SOLID, 1, RGB(220, 220, 220));
                HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                MoveToEx(hdc, rc.right - 1, rc.top + 6, NULL);
                LineTo(hdc, rc.right - 1, rc.bottom - 6);
                SelectObject(hdc, hOldPen);
                DeleteObject(hPen);
            }

            WCHAR szText[MAX_PATH];
            TCITEMW tci = { TCIF_TEXT };
            tci.pszText = szText;
            tci.cchTextMax = ARRAYSIZE(szText);
            TabCtrl_GetItem(hTab, iItem, &tci);

            RECT rcClose = rc;
            rcClose.left = rcClose.right - 22;
            rcClose.top += 3;
            rcClose.bottom -= 3;
            RECT rcText = rc;
            rcText.right = rcClose.left - 4;
            rcText.left += 5;

            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, textColor);
            DrawTextW(hdc, szText, -1, &rcText, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(hTab, &pt);
            TCHITTESTINFO hti = { pt, 0 };

            bool isHovering = (iItem == g_hoveredCloseButtonTab);

            if (isHovering) {
                HBRUSH hHoverBrush = CreateSolidBrush(RGB(232, 17, 35));
                FillRect(hdc, &rcClose, hHoverBrush);
                DeleteObject(hHoverBrush);
                SetTextColor(hdc, RGB(255, 255, 255));
            }
            else {
                SetTextColor(hdc, GetSysColor(COLOR_BTNTEXT));
            }

            HFONT hOldFont = (HFONT)SelectObject(hdc, hMarlettFont);
            DrawTextW(hdc, L"r", 1, &rcClose, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            SelectObject(hdc, hOldFont);

            if (lpdis->itemState & ODS_FOCUS) {
                DrawFocusRect(hdc, &rc);
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
    }
    break;

    case WM_SIZE:
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
        case ID_ACCELERATOR_COPY:
            DoCopyCut(hWnd, false);
            break;
        case ID_ACCELERATOR_CUT:
            DoCopyCut(hWnd, true);
            break;
        case ID_ACCELERATOR_PASTE:
            DoPaste(hWnd);
            break;
        case ID_ACCELERATOR_NEW_TAB:
            AddNewTab(hWnd, pData->currentPath);
            break;
        case ID_ACCELERATOR_CLOSE_TAB:
            CloseTab(hWnd, TabCtrl_GetCurSel(hTab));
            break;
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
                    if (pItem->iItem >= pData->fileData.size()) break;
                    const WIN32_FIND_DATAW& fd = pData->fileData[pItem->iItem];
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
                    std::sort(pData->fileData.begin(), pData->fileData.end(), [&](const WIN32_FIND_DATAW& a, const WIN32_FIND_DATAW& b) {
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
        DeleteObject(hMarlettFont);
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}


// --- タブ管理関数 ---

void AddNewTab(HWND hWnd, LPCWSTR initialPath) {
    auto pData = std::make_unique<ExplorerTabData>();
    if (!pData) return;

    pData->sortColumn = 0;
    pData->sortAscending = true;
    pData->searchTimerId = 0;
    pData->bProgrammaticPathChange = false;
    StringCchCopy(pData->currentPath, MAX_PATH, initialPath);

    int tabIndex = g_tabs.size();

    // ▼▼▼ 変更 ▼▼▼ UpdateLayoutで位置が決まるので、作成時のサイズは仮でOK
    pData->hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | ES_AUTOHSCROLL,
        0, 0, 0, 0,
        hWnd, (HMENU)(EDIT_ID_BASE + tabIndex), GetModuleHandle(NULL), NULL);

    pData->hList = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | LVS_REPORT | LVS_OWNERDATA,
        0, 0, 0, 0,
        hWnd, (HMENU)(LIST_ID_BASE + tabIndex), GetModuleHandle(NULL), NULL);

    if (!pData->hEdit || !pData->hList) return;

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
    if (index < 0 || index >= g_tabs.size()) return;
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
        BOOL show = (i == newIndex);
        ShowWindow(g_tabs[i]->hEdit, show ? SW_SHOW : SW_HIDE);
        ShowWindow(g_tabs[i]->hList, show ? SW_SHOW : SW_HIDE);
    }
    if (newIndex >= 0 && newIndex < g_tabs.size()) {
        SetFocus(g_tabs[newIndex]->hList);
    }
    UpdateLayout(hWnd);
}

void UpdateLayout(HWND hWnd) {
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);

    constexpr int BUTTON_WIDTH = 22;
    constexpr int BUTTON_MARGIN = 4;

    // タブバーの高さを計算 (変更なし)
    int tabHeight = 30;
    int tabCount = TabCtrl_GetItemCount(hTab);
    if (tabCount > 0) {
        RECT rcItem;
        TabCtrl_GetItemRect(hTab, 0, &rcItem);
        tabHeight = rcItem.bottom - rcItem.top + 6;
    }

    int buttonX = 0;
    int buttonY = (tabHeight - BUTTON_WIDTH - 4) / 2 + 2;

    if (tabCount > 0) {
        // --- ▼▼▼ ここから判定ロジックを修正 ▼▼▼ ---

        // 1. スクロールボタンが「現在表示されているか」を確実にチェックする
        HWND hUpDown = FindWindowEx(hTab, NULL, UPDOWN_CLASS, NULL);
        bool scrollButtonsVisible = (hUpDown != NULL && IsWindowVisible(hUpDown));

        // 2. 「+」ボタンを置いた場合に必要な合計幅を計算
        RECT rcLastTab;
        TabCtrl_GetItemRect(hTab, tabCount - 1, &rcLastTab);
        int requiredWidth = rcLastTab.right + BUTTON_MARGIN + BUTTON_WIDTH + BUTTON_MARGIN;

        // 3. 新しい条件でレイアウトを分岐
        if (scrollButtonsVisible || requiredWidth >= rcClient.right) {
            // Case 1: スクロールボタンが「表示されている」または「ボタンを置くと溢れる」場合
            // → 安全な「右端固定モード」にする
            constexpr int RESERVED_AREA_WIDTH = BUTTON_WIDTH + BUTTON_MARGIN * 2;
            int tabWidth = rcClient.right - RESERVED_AREA_WIDTH;
            if (tabWidth < 0) tabWidth = 0;

            MoveWindow(hTab, 0, 0, tabWidth, tabHeight, TRUE);
            buttonX = tabWidth + BUTTON_MARGIN;
        }
        else {
            // Case 2: スクロールボタンが無く、「ボタンを置いてもウィンドウ内に収まる」場合
            // → 「最後のタブの右隣モード」にする
            buttonX = rcLastTab.right + BUTTON_MARGIN;
            MoveWindow(hTab, 0, 0, rcClient.right, tabHeight, TRUE);
        }
    }
    else {
        // タブが一つもない場合 (変更なし)
        buttonX = BUTTON_MARGIN;
        MoveWindow(hTab, 0, 0, rcClient.right, tabHeight, TRUE);
    }

    // 計算したX座標に「+」ボタンを移動
    MoveWindow(hAddButton, buttonX, buttonY, BUTTON_WIDTH, BUTTON_WIDTH, TRUE);
    SetWindowPos(hAddButton, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);

    // --- ▲▲▲ ここまで判定ロジックを修正 ▲▲▲ ---

    // コンテンツ領域のレイアウト (変更なし)
    int curSel = TabCtrl_GetCurSel(hTab);
    if (curSel != -1) {
        ExplorerTabData* pData = GetCurrentTabData();
        if (pData) {
            int contentTop = tabHeight;
            MoveWindow(pData->hEdit, 0, contentTop, rcClient.right, 24, TRUE);
            MoveWindow(pData->hList, 0, contentTop + 24, rcClient.right, rcClient.bottom - contentTop - 24, TRUE);
            UpdateWindow(pData->hList);
        }
    }

    InvalidateRect(hWnd, NULL, TRUE);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPWSTR pCmdLine, int nCmdShow) {
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
        (HBRUSH)(COLOR_WINDOW + 1),
        0,
        szClassName
    };
    RegisterClassW(&wndclass);

    // ▼▼▼ 変更 ▼▼▼ CreateWindowEx から CreateWindow に戻し、WS_EX_COMPOSITED を削除
    HWND hWnd = CreateWindow(
        szClassName, L"タブファイラ",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
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