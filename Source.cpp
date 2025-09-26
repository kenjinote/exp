#define _WIN32_WINNT 0x0A00
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
#include <Aclapi.h>
#include <lm.h>
#include <activeds.h>
#include <sddl.h>
#include <map>
#include <set>
enum WINDOWCOMPOSITIONATTRIB
{
    WCA_USEDARKMODECOLORS = 26
};
struct WINDOWCOMPOSITIONATTRIBDATA
{
    WINDOWCOMPOSITIONATTRIB Attrib;
    PVOID pvData;
    SIZE_T cbData;
};
using fnSetWindowCompositionAttribute = BOOL(WINAPI*)(HWND hWnd, WINDOWCOMPOSITIONATTRIBDATA*);
using fnAllowDarkModeForWindow = bool (WINAPI*)(HWND hWnd, bool allow);
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "pathcch.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "netapi32.lib")
#pragma comment(lib, "activeds.lib")
#pragma comment(lib, "adsiid.lib")
CRITICAL_SECTION g_cacheLock;
std::map<std::wstring, std::wstring> g_sidNameCache;
std::set<std::wstring> g_resolvingSids;
static std::map<HWND, int> g_hoveredItems;
#define WM_APP_OWNER_RESOLVED (WM_APP + 2)
struct OwnerResolveParam {
    HWND hNotifyWnd;
    std::wstring sidString;
    std::wstring accountName;
    std::wstring domainName;
};
struct OwnerResolveResult {
    std::wstring sidString;
    std::wstring displayName;
    bool success;
};
#define IDC_TAB 1000
#define LIST_ID_BASE 2001
#define IDC_ADD_TAB_BUTTON 3001
#define IDC_PATH_EDIT 3002
#define SEARCH_TIMER_ID 1
#define FILTER_TIMER_ID 2
#define WM_APP_FOLDER_SIZE_CALCULATED (WM_APP + 1)
#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif
#ifndef DWMWA_SYSTEM_BACKDROP_TYPE
#define DWMWA_SYSTEM_BACKDROP_TYPE 38
#endif
#ifndef DWMSBT_MAINWINDOW
#define DWMSBT_MAINWINDOW 2
#endif
bool g_isDarkMode = false;
COLORREF g_clrBg;
COLORREF g_clrText;
COLORREF g_clrActiveTab;
COLORREF g_clrAccent;
COLORREF g_clrSeparator;
COLORREF g_clrCloseHoverBg;
COLORREF g_clrCloseHoverText;
COLORREF g_clrCloseText;
COLORREF g_clrSelection;
COLORREF g_clrSelectionText;
COLORREF g_clrHot;
COLORREF g_clrHeaderBg;
static HBRUSH g_hbrBg = NULL;
static HBRUSH g_hbrHeaderBg = NULL;
#ifndef HSIS_SORTEDUP
#define HSIS_SORTEDUP 1
#define HSIS_SORTEDDOWN 2
#endif
#ifndef TCBS_NORMAL
#define TCBS_NORMAL 1
#define TCBS_HOT 2
#endif
#ifndef TABP_CLOSEBUTTON
#define TABP_CLOSEBUTTON 8
#endif
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
enum class SizeCalculationState {
    NotStarted,
    InProgress,
    Completed,
    Error
};
struct FileInfo {
    WIN32_FIND_DATAW findData;
    std::wstring owner;
    std::wstring sidString;
    ULONGLONG calculatedSize;
    SizeCalculationState sizeState;
    int iIcon;
};
struct ExplorerTabData {
    HWND hList;
    WCHAR currentPath[MAX_PATH];
    std::vector<FileInfo> allFileData;
    std::vector<FileInfo> fileData;
    int sortColumn;
    bool sortAscending;
    std::wstring searchString;
    UINT_PTR searchTimerId;
    bool bProgrammaticPathChange;
};
struct FolderSizeParam {
    HWND hNotifyWnd;
    std::wstring folderPath;
};
struct FolderSizeResult {
    std::wstring folderPath;
    ULONGLONG size;
    bool success;
};
WCHAR szClassName[] = L"fai";
static HWND hTab;
static HWND hAddButton;
static HWND hPathEdit;
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
LRESULT CALLBACK PathEditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK ListSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
void ListDirectory(HWND hWnd, ExplorerTabData* pData);
void UpdateSortMark(ExplorerTabData* pData);
void UpdateTabTitle(ExplorerTabData* pData);
void UpdateCaptionButtonsRect(HWND hWnd);
void DrawCaptionButtons(HDC hdc, HWND hWnd);
void FileTimeToString(const FILETIME& ft, WCHAR* buf, size_t bufSize);
void ApplyAddressBarFilter(ExplorerTabData* pData);
DWORD WINAPI CalculateFolderSizeThread(LPVOID lpParam);
void UpdateTheme(HWND hWnd);
LRESULT CALLBACK CustomTabSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
LRESULT CALLBACK TabCtrlSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
int CalculateNameColumnWidth(ExplorerTabData* pData);
LRESULT CALLBACK HeaderSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
ExplorerTabData* GetCurrentTabData() {
    int sel = TabCtrl_GetCurSel(hTab);
    if (sel >= 0 && sel < (int)g_tabs.size()) {
        return g_tabs[sel].get();
    }
    return nullptr;
}
ExplorerTabData* GetTabDataFromChild(HWND hChild) {
    for (const auto& tab : g_tabs) {
        if (tab->hList == hChild) {
            return tab.get();
        }
    }
    return nullptr;
}
std::wstring GetUserDisplayNameFromAD(LPCWSTR samAccountName) {
    std::wstring resultName = samAccountName;
    IADs* pRootDSE = nullptr;
    HRESULT hr = ADsGetObject(L"LDAP://rootDSE", IID_IADs, (void**)&pRootDSE);
    if (SUCCEEDED(hr)) {
        VARIANT varNamingContext;
        VariantInit(&varNamingContext);
        BSTR bstrPropName = SysAllocString(L"defaultNamingContext");
        if (bstrPropName) {
            hr = pRootDSE->Get(bstrPropName, &varNamingContext);
            SysFreeString(bstrPropName);
        }
        else {
            hr = E_OUTOFMEMORY;
        }
        if (SUCCEEDED(hr) && varNamingContext.vt == VT_BSTR) {
            std::wstring ldapPath = L"LDAP://";
            ldapPath += varNamingContext.bstrVal;
            IDirectorySearch* pSearch = nullptr;
            hr = ADsGetObject(ldapPath.c_str(), IID_IDirectorySearch, (void**)&pSearch);
            if (SUCCEEDED(hr)) {
                std::wstring filter = L"(&(objectCategory=person)(objectClass=user)(sAMAccountName=";
                filter += samAccountName;
                filter += L"))";
                LPCWSTR attrs[] = { L"displayName", L"cn" };
                ADS_SEARCH_HANDLE hSearch = NULL;
                hr = pSearch->ExecuteSearch(
                    (LPWSTR)filter.c_str(),
                    (LPWSTR*)attrs,
                    2,
                    &hSearch);
                if (SUCCEEDED(hr)) {
                    if (SUCCEEDED(pSearch->GetFirstRow(hSearch))) {
                        ADS_SEARCH_COLUMN col;
                        if (SUCCEEDED(pSearch->GetColumn(hSearch, (LPWSTR)attrs[0], &col))) {
                            if (col.dwNumValues > 0 && col.pADsValues->dwType == ADSTYPE_CASE_IGNORE_STRING && wcslen(col.pADsValues->CaseIgnoreString) > 0) {
                                resultName = col.pADsValues->CaseIgnoreString;
                            }
                            pSearch->FreeColumn(&col);
                        }
                        if (resultName == samAccountName) {
                            if (SUCCEEDED(pSearch->GetColumn(hSearch, (LPWSTR)attrs[1], &col))) {
                                if (col.dwNumValues > 0 && col.pADsValues->dwType == ADSTYPE_CASE_IGNORE_STRING && wcslen(col.pADsValues->CaseIgnoreString) > 0) {
                                    resultName = col.pADsValues->CaseIgnoreString;
                                }
                                pSearch->FreeColumn(&col);
                            }
                        }
                    }
                    pSearch->CloseSearchHandle(hSearch);
                }
                pSearch->Release();
            }
            VariantClear(&varNamingContext);
        }
        pRootDSE->Release();
    }
    return resultName;
}
DWORD WINAPI ResolveOwnerNameThread(LPVOID lpParam) {
    auto* param = static_cast<OwnerResolveParam*>(lpParam);
    if (!param) return 1;
    HRESULT hrCoInit = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    std::wstring displayName;
    bool success = false;
    if (SUCCEEDED(hrCoInit)) {
        WCHAR szComputerName[MAX_COMPUTERNAME_LENGTH + 1];
        DWORD dwSize = ARRAYSIZE(szComputerName);
        GetComputerNameW(szComputerName, &dwSize);
        if (!param->domainName.empty() && _wcsicmp(param->domainName.c_str(), szComputerName) != 0) {
            displayName = GetUserDisplayNameFromAD(param->accountName.c_str());
            success = (displayName != param->accountName);
        }
        else {
            USER_INFO_2* pUserInfo = nullptr;
            if (NetUserGetInfo(NULL, param->accountName.c_str(), 2, (LPBYTE*)&pUserInfo) == NERR_Success) {
                if (pUserInfo->usri2_full_name && wcslen(pUserInfo->usri2_full_name) > 0) {
                    displayName = pUserInfo->usri2_full_name;
                }
                else {
                    displayName = param->accountName;
                }
                NetApiBufferFree(pUserInfo);
                success = true;
            }
            else {
                displayName = param->accountName;
            }
        }
        CoUninitialize();
    }
    else {
        displayName = param->accountName;
    }
    auto* result = new OwnerResolveResult{ param->sidString, displayName, success };
    PostMessage(param->hNotifyWnd, WM_APP_OWNER_RESOLVED, 0, (LPARAM)result);
    delete param;
    return 0;
}
bool CompareFunction(const FileInfo& a, const FileInfo& b, ExplorerTabData* pData) {
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
        ULARGE_INTEGER sizeA, sizeB;
        if (isDirA) sizeA.QuadPart = a.calculatedSize;
        else { sizeA.LowPart = a.findData.nFileSizeLow; sizeA.HighPart = a.findData.nFileSizeHigh; }
        if (isDirB) sizeB.QuadPart = b.calculatedSize;
        else { sizeB.LowPart = b.findData.nFileSizeLow; sizeB.HighPart = b.findData.nFileSizeHigh; }
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
    case 4:
    {
        std::wstring ownerA, ownerB;
        EnterCriticalSection(&g_cacheLock);
        auto itA = g_sidNameCache.find(a.sidString);
        if (itA != g_sidNameCache.end()) ownerA = itA->second;
        auto itB = g_sidNameCache.find(b.sidString);
        if (itB != g_sidNameCache.end()) ownerB = itB->second;
        LeaveCriticalSection(&g_cacheLock);
        result = ownerA.compare(ownerB);
        break;
    }
    }
    return pData->sortAscending ? (result < 0) : (result > 0);
}
ULONGLONG GetFolderSize(const std::wstring& folderPath) {
    ULONGLONG totalSize = 0;
    WIN32_FIND_DATAW fd;
    std::wstring searchPath = folderPath + L"\\*";
    HANDLE hFind = FindFirstFileW(searchPath.c_str(), &fd);
    if (hFind == INVALID_HANDLE_VALUE) return 0;
    do {
        if (wcscmp(fd.cFileName, L".") != 0 && wcscmp(fd.cFileName, L"..") != 0) {
            std::wstring fullPath = folderPath + L"\\" + fd.cFileName;
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                totalSize += GetFolderSize(fullPath);
            }
            else {
                ULARGE_INTEGER fileSize;
                fileSize.LowPart = fd.nFileSizeLow;
                fileSize.HighPart = fd.nFileSizeHigh;
                totalSize += fileSize.QuadPart;
            }
        }
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);
    return totalSize;
}
DWORD WINAPI CalculateFolderSizeThread(LPVOID lpParam) {
    auto* param = static_cast<FolderSizeParam*>(lpParam);
    if (!param) return 1;
    ULONGLONG size = GetFolderSize(param->folderPath);
    auto* result = new FolderSizeResult{ param->folderPath, size, true };
    PostMessage(param->hNotifyWnd, WM_APP_FOLDER_SIZE_CALCULATED, 0, (LPARAM)result);
    delete param;
    return 0;
}
void ListDirectory(HWND hWnd, ExplorerTabData* pData) {
    pData->allFileData.clear();
    if (PathIsRootW(pData->currentPath) == FALSE) {
        FileInfo up_info = { 0 };
        up_info.findData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        StringCchCopyW(up_info.findData.cFileName, MAX_PATH, L"..");
        up_info.iIcon = -1;
        pData->allFileData.push_back(up_info);
    }
    WIN32_FIND_DATAW fd;
    WCHAR searchPath[MAX_PATH];
    PathCombineW(searchPath, pData->currentPath, L"*");
    HANDLE hFind = FindFirstFileW(searchPath, &fd);
    if (hFind != INVALID_HANDLE_VALUE) {
        do {
            if (wcscmp(fd.cFileName, L".") != 0 && wcscmp(fd.cFileName, L"..") != 0) {
                FileInfo info = { 0 };
                info.findData = fd;
                info.iIcon = -1;
                WCHAR fullPath[MAX_PATH];
                PathCombineW(fullPath, pData->currentPath, fd.cFileName);
                PSID pSidOwner = NULL;
                PSECURITY_DESCRIPTOR pSD = NULL;
                if (GetNamedSecurityInfoW(fullPath, SE_FILE_OBJECT, OWNER_SECURITY_INFORMATION, &pSidOwner, NULL, NULL, NULL, (LPVOID*)&pSD) == ERROR_SUCCESS) {
                    LPWSTR pStringSid = NULL;
                    if (ConvertSidToStringSidW(pSidOwner, &pStringSid)) {
                        info.sidString = pStringSid;
                        LocalFree(pStringSid);
                    }
                    LocalFree(pSD);
                }
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                    info.sizeState = SizeCalculationState::NotStarted;
                    info.calculatedSize = 0;
                }
                pData->allFileData.push_back(info);
            }
        } while (FindNextFileW(hFind, &fd));
        FindClose(hFind);
    }
    std::sort(pData->allFileData.begin(), pData->allFileData.end(), [&](const FileInfo& a, const FileInfo& b) {
        return CompareFunction(a, b, pData);
        });
    for (auto& item : pData->allFileData) {
        if (item.sidString.empty()) continue;
        EnterCriticalSection(&g_cacheLock);
        bool needsResolving = (g_sidNameCache.find(item.sidString) == g_sidNameCache.end() &&
            g_resolvingSids.find(item.sidString) == g_resolvingSids.end());
        if (needsResolving) g_resolvingSids.insert(item.sidString);
        LeaveCriticalSection(&g_cacheLock);
        if (needsResolving) {
            PSID pSid = NULL;
            if (ConvertStringSidToSidW(item.sidString.c_str(), &pSid)) {
                WCHAR szAccountName[256], szDomainName[256];
                DWORD dwAccSize = ARRAYSIZE(szAccountName), dwDomSize = ARRAYSIZE(szDomainName);
                SID_NAME_USE use;
                if (LookupAccountSidW(NULL, pSid, szAccountName, &dwAccSize, szDomainName, &dwDomSize, &use)) {
                    if (use == SidTypeUser) {
                        auto* param = new OwnerResolveParam{ hWnd, item.sidString, szAccountName, szDomainName };
                        QueueUserWorkItem(ResolveOwnerNameThread, param, WT_EXECUTEDEFAULT);
                    }
                    else {
                        EnterCriticalSection(&g_cacheLock);
                        g_sidNameCache[item.sidString] = szAccountName;
                        g_resolvingSids.erase(item.sidString);
                        LeaveCriticalSection(&g_cacheLock);
                    }
                }
                else {
                    EnterCriticalSection(&g_cacheLock);
                    g_sidNameCache[item.sidString] = L"N/A";
                    g_resolvingSids.erase(item.sidString);
                    LeaveCriticalSection(&g_cacheLock);
                }
                LocalFree(pSid);
            }
        }
    }
    for (auto& item : pData->allFileData) {
        if ((item.findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && wcscmp(item.findData.cFileName, L"..") != 0) {
            item.sizeState = SizeCalculationState::InProgress;
            auto* param = new FolderSizeParam;
            param->hNotifyWnd = hWnd;
            WCHAR fullPath[MAX_PATH];
            PathCombineW(fullPath, pData->currentPath, item.findData.cFileName);
            param->folderPath = fullPath;
            QueueUserWorkItem(CalculateFolderSizeThread, param, WT_EXECUTEDEFAULT);
        }
    }
    pData->fileData = pData->allFileData;
    ListView_SetItemCountEx(pData->hList, pData->fileData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(pData->hList, NULL, TRUE);
}
void GetSelectedFilePathsW(ExplorerTabData* pData, std::vector<std::wstring>& paths) {
    paths.clear();
    int iItem = -1;
    while ((iItem = ListView_GetNextItem(pData->hList, iItem, LVNI_SELECTED)) != -1) {
        if (iItem < (int)pData->fileData.size()) {
            const WIN32_FIND_DATAW& fd = pData->fileData[iItem].findData;
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
        if (!isCut) {
            std::wstring tsv_text;
            tsv_text += L"名前\tサイズ\t更新日時\t作成日時\t所有者\r\n";
            int iItem = -1;
            while ((iItem = ListView_GetNextItem(pData->hList, iItem, LVNI_SELECTED)) != -1) {
                if (iItem >= (int)pData->fileData.size()) continue;
                const FileInfo& info = pData->fileData[iItem];
                const WIN32_FIND_DATAW& fd = info.findData;
                if (wcscmp(fd.cFileName, L"..") == 0) continue;
                WCHAR tempBuffer[256];
                tsv_text += fd.cFileName;
                tsv_text += L'\t';
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                    if (wcscmp(fd.cFileName, L"..") != 0) {
                        switch (info.sizeState) {
                        case SizeCalculationState::Completed:
                            StrFormatByteSizeW(info.calculatedSize, tempBuffer, ARRAYSIZE(tempBuffer));
                            tsv_text += tempBuffer;
                            break;
                        case SizeCalculationState::InProgress:
                        case SizeCalculationState::NotStarted:
                            tsv_text += L"計測中...";
                            break;
                        case SizeCalculationState::Error:
                            tsv_text += L"アクセス不可";
                            break;
                        }
                    }
                }
                else {
                    ULARGE_INTEGER sz = { fd.nFileSizeLow, fd.nFileSizeHigh };
                    StrFormatByteSizeW(sz.QuadPart, tempBuffer, ARRAYSIZE(tempBuffer));
                    tsv_text += tempBuffer;
                }
                tsv_text += L'\t';
                FileTimeToString(fd.ftLastWriteTime, tempBuffer, ARRAYSIZE(tempBuffer));
                tsv_text += tempBuffer;
                tsv_text += L'\t';
                FileTimeToString(fd.ftCreationTime, tempBuffer, ARRAYSIZE(tempBuffer));
                tsv_text += tempBuffer;
                tsv_text += L'\t';
                if (!info.sidString.empty()) {
                    EnterCriticalSection(&g_cacheLock);
                    auto it = g_sidNameCache.find(info.sidString);
                    if (it != g_sidNameCache.end()) {
                        tsv_text += it->second;
                    }
                    else {
                        tsv_text += L"取得中...";
                    }
                    LeaveCriticalSection(&g_cacheLock);
                }
                tsv_text += L"\r\n";
            }
            if (!tsv_text.empty()) {
                HGLOBAL hText = GlobalAlloc(GMEM_MOVEABLE, (tsv_text.length() + 1) * sizeof(WCHAR));
                if (hText) {
                    WCHAR* pText = (WCHAR*)GlobalLock(hText);
                    if (pText) {
                        StringCchCopyW(pText, tsv_text.length() + 1, tsv_text.c_str());
                        GlobalUnlock(hText);
                        SetClipboardData(CF_UNICODETEXT, hText);
                    }
                    else {
                        GlobalFree(hText);
                    }
                }
            }
        }
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
    ListDirectory(hWnd, pData);
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
    ListDirectory(hWnd, pData);
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
                const WIN32_FIND_DATAW& fd = pData->fileData[currentSel].findData;
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
                    ListDirectory(hWnd, pData);
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
        SetWindowTextW(hPathEdit, pData->currentPath);
        ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED);
        ListDirectory(hWnd, pData);
        UpdateSortMark(pData);
        for (size_t i = 0; i < pData->fileData.size(); ++i) {
            if (wcscmp(pData->fileData[i].findData.cFileName, previousFolderName) == 0) {
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
    const WIN32_FIND_DATAW& fd = pData->fileData[sel].findData;
    if (wcscmp(fd.cFileName, L"..") == 0) {
        NavigateUpAndSelect(GetParent(pData->hList));
        return;
    }
    if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
        WCHAR newPath[MAX_PATH];
        PathCombineW(newPath, pData->currentPath, fd.cFileName);
        StringCchCopyW(pData->currentPath, MAX_PATH, newPath);
        UpdateTabTitle(pData);
        SetWindowTextW(hPathEdit, pData->currentPath);
        ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED);
        ListDirectory(GetParent(pData->hList), pData);
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
LRESULT CALLBACK PathEditSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    ExplorerTabData* pData = GetCurrentTabData();
    if (!pData) return DefSubclassProc(hWnd, uMsg, wParam, lParam);
    switch (uMsg)
    {
    case WM_KEYDOWN:
        if (wParam == VK_RETURN)
        {
            WCHAR buffer[MAX_PATH];
            GetWindowTextW(hWnd, buffer, MAX_PATH);
            WCHAR resolvedPath[MAX_PATH];
            if (PathIsRelativeW(buffer)) {
                PathCchCombine(resolvedPath, MAX_PATH, pData->currentPath, buffer);
            }
            else {
                StringCchCopyW(resolvedPath, MAX_PATH, buffer);
            }
            if (PathIsDirectoryW(resolvedPath)) {
                PathCchCanonicalize(pData->currentPath, MAX_PATH, resolvedPath);
                UpdateTabTitle(pData);
                pData->bProgrammaticPathChange = true;
                SetWindowTextW(hWnd, pData->currentPath);
                pData->bProgrammaticPathChange = false;
                ListDirectory(GetParent(hWnd), pData);
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
                ApplyAddressBarFilter(pData);
            }
            return 0;
        }
        else if (wParam == VK_ESCAPE)
        {
            pData->bProgrammaticPathChange = true;
            SetWindowTextW(hWnd, pData->currentPath);
            pData->bProgrammaticPathChange = false;
            ApplyAddressBarFilter(pData);
            SendMessageW(hWnd, EM_SETSEL, 0, -1);
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
    case WM_KILLFOCUS:
    {
        if (GetWindowTextLengthW(hWnd) == 0)
        {
            pData->bProgrammaticPathChange = true;
            SetWindowTextW(hWnd, pData->currentPath);
            pData->bProgrammaticPathChange = false;
        }
        break;
    }
    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, PathEditSubclassProc, uIdSubclass);
        break;
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}
LRESULT CALLBACK CustomTabSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    static bool bTracking = false;
    static bool g_isDragging = false;
    static int g_draggedTab = -1;
    static POINT g_ptDragStart;
    static int g_markIndex = -1;
    static BOOL g_markAfter = FALSE;
    switch (uMsg)
    {
    case WM_PAINT:
    {
        LRESULT result = DefSubclassProc(hWnd, uMsg, wParam, lParam);
        if (g_isDragging && g_markIndex != -1) {
            bool isAtStart = (!g_markAfter && g_markIndex == g_draggedTab) || (g_markAfter && g_markIndex == g_draggedTab - 1);
            bool isAtEnd = (g_markAfter && g_markIndex == g_draggedTab) || (!g_markAfter && g_markIndex == g_draggedTab + 1);
            if (!(isAtStart || isAtEnd)) {
                HDC hdc = GetDC(hWnd);
                int x;
                RECT rcCurrent;
                TabCtrl_GetItemRect(hWnd, g_markIndex, &rcCurrent);
                if (g_markAfter) {
                    int nextIndex = g_markIndex + 1;
                    if (nextIndex < TabCtrl_GetItemCount(hWnd)) {
                        RECT rcNext;
                        TabCtrl_GetItemRect(hWnd, nextIndex, &rcNext);
                        x = rcCurrent.right + (rcNext.left - rcCurrent.right) / 2;
                    }
                    else {
                        x = rcCurrent.right;
                    }
                }
                else {
                    if (g_markIndex > 0) {
                        RECT rcPrev;
                        TabCtrl_GetItemRect(hWnd, g_markIndex - 1, &rcPrev);
                        x = rcPrev.right + (rcCurrent.left - rcPrev.right) / 2;
                    }
                    else {
                        x = rcCurrent.left;
                    }
                }
                LOGBRUSH lb;
                lb.lbStyle = BS_SOLID;
                lb.lbColor = g_clrText;
                HPEN hPen = ExtCreatePen(PS_GEOMETRIC | PS_SOLID | PS_ENDCAP_SQUARE, 3, &lb, 0, NULL);
                HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                MoveToEx(hdc, x, rcCurrent.top + 2, NULL);
                LineTo(hdc, x, rcCurrent.bottom - 2);
                SelectObject(hdc, hOldPen);
                DeleteObject(hPen);
                ReleaseDC(hWnd, hdc);
            }
        }
        return result;
    }
    case WM_LBUTTONDOWN:
    {
        TCHITTESTINFO hti;
        g_ptDragStart.x = hti.pt.x = GET_X_LPARAM(lParam);
        g_ptDragStart.y = hti.pt.y = GET_Y_LPARAM(lParam);
        g_markIndex = -1;
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
            else
            {
                g_draggedTab = iTab;
                g_isDragging = true;
                SetCapture(hWnd);
                return 0;
            }
        }
        break;
    }
    case WM_LBUTTONUP:
    {
        if (g_isDragging) {
            ReleaseCapture();
            g_isDragging = false;
            g_markIndex = -1;
            InvalidateRect(hWnd, NULL, TRUE);
            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(hWnd, &pt);
            if (abs(pt.x - g_ptDragStart.x) < GetSystemMetrics(SM_CXDRAG) &&
                abs(pt.y - g_ptDragStart.y) < GetSystemMetrics(SM_CYDRAG))
            {
                SwitchTab(GetParent(hWnd), g_draggedTab);
                g_draggedTab = -1;
                return 0;
            }
            TCHITTESTINFO hti = { 0 };
            hti.pt = pt;
            if (TabCtrl_GetItemCount(hWnd) > 0) {
                RECT rcFirstTab;
                TabCtrl_GetItemRect(hWnd, 0, &rcFirstTab);
                hti.pt.y = rcFirstTab.top + (rcFirstTab.bottom - rcFirstTab.top) / 2;
            }
            int dropIndex = TabCtrl_HitTest(hWnd, &hti);
            int finalDropIndex = -1;
            if (dropIndex != -1) {
                RECT rcItem;
                TabCtrl_GetItemRect(hWnd, dropIndex, &rcItem);
                BOOL fAfter = (pt.x > rcItem.left + (rcItem.right - rcItem.left) / 2);
                finalDropIndex = dropIndex;
                if (fAfter) {
                    finalDropIndex++;
                }
            }
            else {
                int tabCount = TabCtrl_GetItemCount(hWnd);
                if (tabCount > 0) {
                    RECT rcFirst, rcLast;
                    TabCtrl_GetItemRect(hWnd, 0, &rcFirst);
                    TabCtrl_GetItemRect(hWnd, tabCount - 1, &rcLast);
                    if (pt.x < rcFirst.left) {
                        finalDropIndex = 0;
                    }
                    else if (pt.x > rcLast.right) {
                        finalDropIndex = tabCount;
                    }
                    else {
                        g_draggedTab = -1;
                        return 0;
                    }
                }
            }
            if (g_draggedTab != -1 && finalDropIndex != g_draggedTab && finalDropIndex != g_draggedTab + 1) {
                if (finalDropIndex > g_draggedTab) {
                    finalDropIndex--;
                }
                if (finalDropIndex >= 0 && finalDropIndex <= (int)g_tabs.size() && finalDropIndex != g_draggedTab) {
                    auto draggedData = std::move(g_tabs[g_draggedTab]);
                    g_tabs.erase(g_tabs.begin() + g_draggedTab);
                    g_tabs.insert(g_tabs.begin() + finalDropIndex, std::move(draggedData));
                    WCHAR szText[MAX_PATH];
                    TCITEMW tci = { TCIF_TEXT };
                    tci.pszText = szText;
                    tci.cchTextMax = MAX_PATH;
                    TabCtrl_GetItem(hWnd, g_draggedTab, &tci);
                    TabCtrl_DeleteItem(hWnd, g_draggedTab);
                    TabCtrl_InsertItem(hWnd, finalDropIndex, &tci);
                    SwitchTab(GetParent(hWnd), finalDropIndex);
                }
            }
            g_draggedTab = -1;
            return 0;
        }
        break;
    }
    case WM_MOUSEMOVE:
    {
        if (g_isDragging) {
            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(hWnd, &pt);
            if (abs(pt.x - g_ptDragStart.x) < GetSystemMetrics(SM_CXDRAG) &&
                abs(pt.y - g_ptDragStart.y) < GetSystemMetrics(SM_CYDRAG)) {
                return 0;
            }
            TCHITTESTINFO hti = { 0 };
            hti.pt = pt;
            if (TabCtrl_GetItemCount(hWnd) > 0) {
                RECT rcFirstTab;
                TabCtrl_GetItemRect(hWnd, 0, &rcFirstTab);
                hti.pt.y = rcFirstTab.top + (rcFirstTab.bottom - rcFirstTab.top) / 2;
            }
            int dropIndex = TabCtrl_HitTest(hWnd, &hti);
            int newMarkIndex = -1;
            BOOL newMarkAfter = FALSE;
            if (dropIndex != -1) {
                newMarkIndex = dropIndex;
                RECT rcItem;
                TabCtrl_GetItemRect(hWnd, dropIndex, &rcItem);
                newMarkAfter = (pt.x > rcItem.left + (rcItem.right - rcItem.left) / 2);
            }
            else {
                int tabCount = TabCtrl_GetItemCount(hWnd);
                if (tabCount > 0) {
                    RECT rcFirst, rcLast;
                    TabCtrl_GetItemRect(hWnd, 0, &rcFirst);
                    TabCtrl_GetItemRect(hWnd, tabCount - 1, &rcLast);
                    if (pt.x < rcFirst.left) {
                        newMarkIndex = 0;
                        newMarkAfter = FALSE;
                    }
                    else if (pt.x > rcLast.right) {
                        newMarkIndex = tabCount - 1;
                        newMarkAfter = TRUE;
                    }
                }
            }
            if (g_markIndex != newMarkIndex || g_markAfter != newMarkAfter) {
                g_markIndex = newMarkIndex;
                g_markAfter = newMarkAfter;
                InvalidateRect(hWnd, NULL, TRUE);
            }
            return 0;
        }
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
    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, CustomTabSubclassProc, uIdSubclass);
        break;
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}
void PerformIncrementalSearch(ExplorerTabData* pData) {
    if (pData->searchString.empty()) return;
    int currentIndex = ListView_GetNextItem(pData->hList, -1, LVNI_FOCUSED);
    int startIndex = (currentIndex == -1) ? 0 : currentIndex;
    for (size_t i = startIndex + 1; i < pData->fileData.size(); ++i) {
        if (StrCmpNIW(pData->fileData[i].findData.cFileName, pData->searchString.c_str(), pData->searchString.length()) == 0) {
            ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_SetItemState(pData->hList, (int)i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_FOCUSED | LVIS_SELECTED);
            ListView_EnsureVisible(pData->hList, (int)i, FALSE);
            return;
        }
    }
    for (int i = 0; i <= startIndex; ++i) {
        if (StrCmpNIW(pData->fileData[i].findData.cFileName, pData->searchString.c_str(), pData->searchString.length()) == 0) {
            ListView_SetItemState(pData->hList, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
            ListView_SetItemState(pData->hList, i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_SELECTED);
            ListView_EnsureVisible(pData->hList, i, FALSE);
            return;
        }
    }
}
LRESULT CALLBACK ListSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    ExplorerTabData* pData = GetTabDataFromChild(hWnd);
    if (!pData) return DefSubclassProc(hWnd, uMsg, wParam, lParam);
    switch (uMsg) {
    case WM_NCHITTEST:
    {
        POINT ptScreen = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        RECT rcWindow;
        GetWindowRect(hWnd, &rcWindow);
        constexpr int BORDER_WIDTH = 8;
        if (ptScreen.x >= rcWindow.left && ptScreen.x < rcWindow.left + BORDER_WIDTH) {
            return HTTRANSPARENT;
        }
        if (ptScreen.x < rcWindow.right && ptScreen.x >= rcWindow.right - BORDER_WIDTH) {
            return HTTRANSPARENT;
        }
        if (ptScreen.y < rcWindow.bottom && ptScreen.y >= rcWindow.bottom - BORDER_WIDTH) {
            return HTTRANSPARENT;
        }
        return DefSubclassProc(hWnd, uMsg, wParam, lParam);
    }
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
void ApplyAddressBarFilter(ExplorerTabData* pData) {
    WCHAR filterText[MAX_PATH];
    GetWindowTextW(hPathEdit, filterText, MAX_PATH);
    if (wcslen(filterText) > 0 && (wcschr(filterText, L'\\') || wcschr(filterText, L'/') || wcschr(filterText, L':'))) {
        if (pData->fileData.size() != pData->allFileData.size()) {
            pData->fileData = pData->allFileData;
            ListView_SetItemCountEx(pData->hList, pData->fileData.size(), LVSICF_NOINVALIDATEALL);
            InvalidateRect(pData->hList, NULL, TRUE);
        }
        return;
    }
    if (wcslen(filterText) == 0) {
        if (pData->fileData.size() != pData->allFileData.size()) {
            pData->fileData = pData->allFileData;
            ListView_SetItemCountEx(pData->hList, pData->fileData.size(), LVSICF_NOINVALIDATEALL);
            InvalidateRect(pData->hList, NULL, TRUE);
        }
        return;
    }
    pData->fileData.clear();
    for (const auto& fileInfo : pData->allFileData) {
        if (StrStrIW(fileInfo.findData.cFileName, filterText)) {
            pData->fileData.push_back(fileInfo);
        }
    }
    ListView_SetItemCountEx(pData->hList, pData->fileData.size(), LVSICF_NOINVALIDATEALL);
    InvalidateRect(pData->hList, NULL, TRUE);
}
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_APP_OWNER_RESOLVED:
    {
        auto* result = reinterpret_cast<OwnerResolveResult*>(lParam);
        if (result) {
            EnterCriticalSection(&g_cacheLock);
            g_sidNameCache[result->sidString] = result->displayName;
            g_resolvingSids.erase(result->sidString);
            LeaveCriticalSection(&g_cacheLock);
            ExplorerTabData* pData = GetCurrentTabData();
            if (pData) { InvalidateRect(pData->hList, NULL, FALSE); }
            delete result;
        }
        return 0;
    }
    case WM_APP_FOLDER_SIZE_CALCULATED:
    {
        auto* result = reinterpret_cast<FolderSizeResult*>(lParam);
        if (result) {
            ExplorerTabData* pData = GetCurrentTabData();
            if (pData && PathIsPrefixW(pData->currentPath, result->folderPath.c_str())) {
                for (auto& item : pData->allFileData) {
                    WCHAR fullPath[MAX_PATH];
                    PathCombineW(fullPath, pData->currentPath, item.findData.cFileName);
                    if (result->folderPath == fullPath) {
                        item.calculatedSize = result->size;
                        item.sizeState = result->success ? SizeCalculationState::Completed : SizeCalculationState::Error;
                        break;
                    }
                }
                for (int i = 0; i < (int)pData->fileData.size(); ++i) {
                    WCHAR fullPath[MAX_PATH];
                    PathCombineW(fullPath, pData->currentPath, pData->fileData[i].findData.cFileName);
                    if (result->folderPath == fullPath) {
                        pData->fileData[i].calculatedSize = result->size;
                        pData->fileData[i].sizeState = result->success ? SizeCalculationState::Completed : SizeCalculationState::Error;
                        ListView_RedrawItems(pData->hList, i, i);
                        break;
                    }
                }
            }
            delete result;
        }
        return 0;
    }
    case WM_SETTINGCHANGE:
    {
        if (lParam != 0 && wcscmp((LPCWSTR)lParam, L"ImmersiveColorSet") == 0)
        {
            UpdateTheme(hWnd);
        }
        return 0;
    }
    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT:
    {
        if ((HWND)lParam == hPathEdit) {
            HDC hdc = (HDC)wParam;
            SetTextColor(hdc, g_clrText);
            SetBkColor(hdc, g_clrBg);
            return (INT_PTR)g_hbrBg;
        }
        break;
    }
    case WM_NCACTIVATE:
        return 1;
    case WM_NCPAINT:
        return 0;
    case WM_NCCALCSIZE:
        return 0;
    case WM_NCHITTEST:
    {
        POINT ptScreen = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        POINT ptClient = ptScreen;
        ScreenToClient(hWnd, &ptClient);
        if (PtInRect(&g_rcMinButton, ptClient) ||
            PtInRect(&g_rcMaxButton, ptClient) ||
            PtInRect(&g_rcCloseButton, ptClient)) {
            return HTCLIENT;
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
    case WM_LBUTTONDOWN:
    {
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        CaptionButtonState pressed = CaptionButtonState::None;
        if (PtInRect(&g_rcMinButton, pt)) pressed = CaptionButtonState::Min;
        else if (PtInRect(&g_rcMaxButton, pt)) pressed = CaptionButtonState::Max;
        else if (PtInRect(&g_rcCloseButton, pt)) pressed = CaptionButtonState::Close;
        if (pressed != CaptionButtonState::None) {
            g_pressedButton = pressed;
            SetCapture(hWnd);
            InvalidateRect(hWnd, NULL, FALSE);
            return 0;
        }
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    case WM_LBUTTONUP:
    {
        if (g_pressedButton != CaptionButtonState::None) {
            CaptionButtonState pressedButton = g_pressedButton;
            g_pressedButton = CaptionButtonState::None;
            ReleaseCapture();
            if (g_hoveredButton == pressedButton) {
                switch (pressedButton) {
                case CaptionButtonState::Min:
                    PostMessage(hWnd, WM_SYSCOMMAND, SC_MINIMIZE, 0);
                    break;
                case CaptionButtonState::Max:
                    PostMessage(hWnd, WM_SYSCOMMAND, IsZoomed(hWnd) ? SC_RESTORE : SC_MAXIMIZE, 0);
                    break;
                case CaptionButtonState::Close:
                    PostMessage(hWnd, WM_SYSCOMMAND, SC_CLOSE, 0);
                    break;
                }
            }
            InvalidateRect(hWnd, NULL, FALSE);
            return 0;
        }
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    case WM_CAPTURECHANGED:
    {
        if ((HWND)lParam != hWnd) {
            g_pressedButton = CaptionButtonState::None;
            InvalidateRect(hWnd, NULL, TRUE);
        }
        return 0;
    }
    case WM_CREATE:
    {
        hFont = CreateFontW(20, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET, 0, 0, 0, 0, L"Segoe UI");
        MARGINS margins = { 0, 0, 1, 0 };
        DwmExtendFrameIntoClientArea(hWnd, &margins);
        SetWindowPos(hWnd, NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOOWNERZORDER);
        hTab = CreateWindowW(WC_TABCONTROLW, 0, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TCS_OWNERDRAWFIXED,
            0, 0, 0, 0, hWnd, (HMENU)IDC_TAB, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
        SendMessage(hTab, WM_SETFONT, (WPARAM)hFont, TRUE);
        TabCtrl_SetPadding(hTab, 18, 5);
        SetWindowSubclass(hTab, CustomTabSubclassProc, 0, 0);
        SetWindowSubclass(hTab, TabCtrlSubclassProc, 0, 0);
        hAddButton = CreateWindowW(L"BUTTON", L"+", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            0, 0, 24, 24, hWnd, (HMENU)IDC_ADD_TAB_BUTTON, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
        hPathEdit = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_BORDER,
            0, 0, 0, 0, hWnd, (HMENU)IDC_PATH_EDIT, ((LPCREATESTRUCT)lParam)->hInstance, NULL);
        SendMessage(hPathEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
        SetWindowSubclass(hPathEdit, PathEditSubclassProc, 0, 0);
        SHAutoComplete(hPathEdit, SHACF_FILESYS_DIRS | SHACF_AUTOSUGGEST_FORCE_ON | SHACF_AUTOAPPEND_FORCE_ON);
        LOGFONTW lf = { 0 };
        GetObject(hFont, sizeof(lf), &lf);
        lf.lfHeight = 22;
        lf.lfWeight = FW_LIGHT;
        HFONT hAddButtonFont = CreateFontIndirectW(&lf);
        SendMessage(hAddButton, WM_SETFONT, (WPARAM)hAddButtonFont, TRUE);
        SHFILEINFO sfi = { 0 };
        hImgList = (HIMAGELIST)SHGetFileInfoW(
            L"dummy",
            FILE_ATTRIBUTE_NORMAL,
            &sfi,
            sizeof(sfi),
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES
        );
        UpdateTheme(hWnd);
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
        if (g_hoveredButton != CaptionButtonState::None) {
            g_hoveredButton = CaptionButtonState::None;
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
        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        RECT rcTitleBar = { 0, 0, rcClient.right, 32 };
        HBRUSH hBrush = CreateSolidBrush(g_clrBg);
        FillRect(hdc, &rcTitleBar, hBrush);
        DeleteObject(hBrush);
        DrawCaptionButtons(hdc, hWnd);
        EndPaint(hWnd, &ps);
        return 0;
    }
    case WM_DRAWITEM:
    {
        LPDRAWITEMSTRUCT lpdis = (LPDRAWITEMSTRUCT)lParam;
        if (lpdis->CtlID == IDC_ADD_TAB_BUTTON) {
            HDC hdc = lpdis->hDC;
            RECT rc = lpdis->rcItem;
            HBRUSH hBrush = CreateSolidBrush(g_clrBg);
            FillRect(hdc, &rc, hBrush);
            DeleteObject(hBrush);
            if (lpdis->itemState & ODS_HOTLIGHT) {
                COLORREF hoverBg = g_isDarkMode ? RGB(80, 80, 80) : RGB(229, 243, 255);
                HBRUSH hFrameBrush = CreateSolidBrush(hoverBg);
                FillRect(hdc, &rc, hFrameBrush);
                DeleteObject(hFrameBrush);
                HPEN hPen = CreatePen(PS_SOLID, 1, g_clrAccent);
                HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
                SelectObject(hdc, hOldPen);
                DeleteObject(hPen);
            }
            if (lpdis->itemState & ODS_SELECTED) {
                COLORREF pushedBg = g_isDarkMode ? RGB(100, 100, 100) : RGB(204, 232, 255);
                HBRUSH hFrameBrush = CreateSolidBrush(pushedBg);
                FillRect(hdc, &rc, hFrameBrush);
                DeleteObject(hFrameBrush);
            }
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, g_clrText);
            DrawTextW(hdc, L"+", -1, &rc, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
            return TRUE;
        }
        if (lpdis->CtlID == IDC_TAB) {
            HDC hdc = lpdis->hDC;
            RECT rcItem = lpdis->rcItem;
            int iItem = lpdis->itemID;
            int iSel = TabCtrl_GetCurSel(hTab);
            bool isSelected = (iItem == iSel);
            COLORREF bgColor = isSelected ? g_clrActiveTab : g_clrBg;
            HBRUSH hBrush = CreateSolidBrush(bgColor);
            FillRect(hdc, &rcItem, hBrush);
            DeleteObject(hBrush);
            if (isSelected) {
                RECT rcIndicator = { rcItem.left, rcItem.bottom - 2, rcItem.right, rcItem.bottom };
                HBRUSH hAccentBrush = CreateSolidBrush(g_clrAccent);
                FillRect(hdc, &rcIndicator, hAccentBrush);
                DeleteObject(hAccentBrush);
            }
            HPEN hPen = CreatePen(PS_SOLID, 1, g_clrSeparator);
            HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
            MoveToEx(hdc, rcItem.right - 1, rcItem.top + 4, NULL);
            LineTo(hdc, rcItem.right - 1, rcItem.bottom - 4);
            SelectObject(hdc, hOldPen);
            DeleteObject(hPen);
            WCHAR szText[MAX_PATH];
            TCITEMW tci = { TCIF_TEXT, 0, 0, szText, ARRAYSIZE(szText) };
            TabCtrl_GetItem(hTab, iItem, &tci);
            RECT rcClose = { rcItem.right - 22, rcItem.top + (rcItem.bottom - rcItem.top - 16) / 2, rcItem.right - 2, 0 };
            rcClose.bottom = rcClose.top + 16;
            RECT rcText = { rcItem.left + 8, rcItem.top, rcClose.left, rcItem.bottom };
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, g_clrText);
            DrawTextW(hdc, szText, -1, &rcText, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            if (iItem == g_hoveredCloseButtonTab) {
                HBRUSH hCloseBrush = CreateSolidBrush(g_clrCloseHoverBg);
                FillRect(hdc, &rcClose, hCloseBrush);
                DeleteObject(hCloseBrush);
                SetTextColor(hdc, g_clrCloseHoverText);
            }
            else {
                SetTextColor(hdc, g_clrCloseText);
            }
            DrawTextW(hdc, L"✕", -1, &rcClose, DT_CENTER | DT_SINGLELINE | DT_VCENTER);
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
        else if (wParam == FILTER_TIMER_ID) {
            KillTimer(hWnd, FILTER_TIMER_ID);
            ExplorerTabData* pData = GetCurrentTabData();
            if (pData) {
                ApplyAddressBarFilter(pData);
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
        if (HIWORD(wParam) == EN_CHANGE && (HWND)lParam == hPathEdit) {
            ExplorerTabData* pData = GetCurrentTabData();
            if (pData && !pData->bProgrammaticPathChange) {
                SetTimer(hWnd, FILTER_TIMER_ID, 200, nullptr);
            }
            return 0;
        }
        ExplorerTabData* pData = GetCurrentTabData();
        if (!pData && LOWORD(wParam) != IDC_ADD_TAB_BUTTON) break;
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
            SetFocus(hPathEdit);
            SendMessage(hPathEdit, EM_SETSEL, 0, -1);
            break;
        case ID_ACCELERATOR_TAB:
        {
            HWND hFocus = GetFocus();
            if (hFocus == hPathEdit) SetFocus(pData->hList);
            else if (hFocus == pData->hList) {
                SetFocus(hTab);
            }
            else if (hFocus == hTab) {
                SetFocus(hAddButton);
            }
            else {
                SetFocus(hPathEdit);
                SendMessage(hPathEdit, EM_SETSEL, 0, -1);
            }
        }
        break;
        case ID_ACCELERATOR_REFRESH:
            ListDirectory(hWnd, pData);
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
        if (lpnmh->code == HDN_DIVIDERDBLCLICK) {
            LPNMHEADER pnmhdr = (LPNMHEADER)lParam;
            ExplorerTabData* pData = nullptr;
            for (const auto& tab : g_tabs) {
                if (ListView_GetHeader(tab->hList) == lpnmh->hwndFrom) {
                    pData = tab.get();
                    break;
                }
            }
            if (pData && pnmhdr->iItem == 0) {
                int maxWidth = CalculateNameColumnWidth(pData);
                LVCOLUMN lvc = { LVCF_WIDTH };
                lvc.cx = maxWidth;
                ListView_SetColumn(pData->hList, 0, &lvc);
                return 0;
            }
        }
        if (lpnmh->idFrom == IDC_TAB) {
            if (lpnmh->code == TCN_SELCHANGE) {
                SwitchTab(hWnd, TabCtrl_GetCurSel(hTab));
            }
        }
        else {
            ExplorerTabData* pData = GetTabDataFromChild(lpnmh->hwndFrom);
            if (pData) {
                switch (lpnmh->code) {
                case NM_CUSTOMDRAW:
                {
                    LPNMLVCUSTOMDRAW lplvcd = (LPNMLVCUSTOMDRAW)lParam;
                    switch (lplvcd->nmcd.dwDrawStage)
                    {
                    case CDDS_PREPAINT:
                        return CDRF_NOTIFYITEMDRAW;
                    case CDDS_ITEMPREPAINT:
                    {
                        HDC hdc = lplvcd->nmcd.hdc;
                        HWND hListView = lplvcd->nmcd.hdr.hwndFrom;
                        int iItem = (int)lplvcd->nmcd.dwItemSpec;
                        UINT uItemState = lplvcd->nmcd.uItemState;
                        ExplorerTabData* pData = GetTabDataFromChild(hListView);
                        if (!pData || iItem >= (int)pData->fileData.size()) {
                            return CDRF_DODEFAULT;
                        }
                        auto it = g_hoveredItems.find(hListView);
                        int hoveredItem = (it != g_hoveredItems.end()) ? it->second : -1;
                        bool isHovered = (hoveredItem == iItem);
                        COLORREF clrText, clrTextBk;
                        if (uItemState & CDIS_SELECTED) {
                            clrText = g_clrSelectionText;
                            clrTextBk = g_clrSelection;
                        }
                        else if (isHovered) {
                            clrText = g_clrText;
                            clrTextBk = g_clrHot;
                        }
                        else {
                            clrText = g_clrText;
                            clrTextBk = g_clrBg;
                        }
                        RECT rcItem;
                        ListView_GetItemRect(hListView, iItem, &rcItem, LVIR_BOUNDS);
                        HBRUSH hBrush = CreateSolidBrush(clrTextBk);
                        FillRect(hdc, &rcItem, hBrush);
                        DeleteObject(hBrush);
                        SetTextColor(hdc, clrText);
                        SetBkMode(hdc, TRANSPARENT);
                        const FileInfo& info = pData->fileData[iItem];
                        const WIN32_FIND_DATAW& fd = info.findData;
                        HWND hHeader = ListView_GetHeader(hListView);
                        int columnCount = Header_GetItemCount(hHeader);
                        std::vector<int> columnPositions;
                        columnPositions.push_back(0);
                        for (int i = 0; i < columnCount; i++) {
                            RECT rcHeader;
                            Header_GetItemRect(hHeader, i, &rcHeader);
                            columnPositions.push_back(rcHeader.right);
                        }
                        for (int iSubItem = 0; iSubItem < columnCount; iSubItem++) {
                            WCHAR szText[MAX_PATH] = { 0 };
                            switch (iSubItem) {
                            case 0:
                                StringCchCopy(szText, MAX_PATH, fd.cFileName);
                                break;
                            case 1:
                                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                                    if (wcscmp(fd.cFileName, L"..") == 0) {
                                        szText[0] = L'\0';
                                    }
                                    else {
                                        switch (info.sizeState) {
                                        case SizeCalculationState::InProgress:
                                        case SizeCalculationState::NotStarted:
                                            StringCchCopy(szText, MAX_PATH, L"計測中...");
                                            break;
                                        case SizeCalculationState::Completed:
                                            StrFormatByteSizeW(info.calculatedSize, szText, MAX_PATH);
                                            break;
                                        case SizeCalculationState::Error:
                                            StringCchCopy(szText, MAX_PATH, L"アクセス不可");
                                            break;
                                        }
                                    }
                                }
                                else {
                                    ULARGE_INTEGER sz = { fd.nFileSizeLow, fd.nFileSizeHigh };
                                    StrFormatByteSizeW(sz.QuadPart, szText, MAX_PATH);
                                }
                                break;
                            case 2:
                            case 3:
                                if (wcscmp(fd.cFileName, L"..") != 0) {
                                    const FILETIME& ft = (iSubItem == 2) ? fd.ftLastWriteTime : fd.ftCreationTime;
                                    FileTimeToString(ft, szText, MAX_PATH);
                                }
                                else {
                                    szText[0] = L'\0';
                                }
                                break;
                            case 4:
                                if (!info.sidString.empty()) {
                                    EnterCriticalSection(&g_cacheLock);
                                    auto it = g_sidNameCache.find(info.sidString);
                                    if (it != g_sidNameCache.end()) {
                                        StringCchCopy(szText, MAX_PATH, it->second.c_str());
                                    }
                                    else {
                                        StringCchCopy(szText, MAX_PATH, L"取得中...");
                                    }
                                    LeaveCriticalSection(&g_cacheLock);
                                }
                                else {
                                    StringCchCopy(szText, MAX_PATH, info.owner.c_str());
                                }
                                break;
                            }
                            RECT rcSubItem = rcItem;
                            rcSubItem.left = columnPositions[iSubItem];
                            rcSubItem.right = columnPositions[iSubItem + 1];
                            if (iSubItem == 0) {
                                RECT rcIcon;
                                ListView_GetSubItemRect(hListView, iItem, iSubItem, LVIR_ICON, &rcIcon);
                                if (info.iIcon == -1) {
                                    WCHAR fullpath[MAX_PATH];
                                    PathCombineW(fullpath, pData->currentPath, fd.cFileName);
                                    SHFILEINFO sfi = { 0 };
                                    SHGetFileInfoW(fullpath,
                                        fd.dwFileAttributes,
                                        &sfi,
                                        sizeof(sfi),
                                        SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
                                    const_cast<FileInfo&>(info).iIcon = sfi.iIcon;
                                    auto it = std::find_if(pData->allFileData.begin(), pData->allFileData.end(),
                                        [&](const FileInfo& master_info) {
                                            return wcscmp(master_info.findData.cFileName, info.findData.cFileName) == 0;
                                        });
                                    if (it != pData->allFileData.end()) {
                                        it->iIcon = sfi.iIcon;
                                    }
                                }
                                int iIcon = info.iIcon;
                                if (iIcon != -1 && hImgList) {
                                    ImageList_Draw(hImgList, iIcon, hdc, rcIcon.left, rcIcon.top, ILD_TRANSPARENT);
                                }
                                RECT rcText = rcSubItem;
                                rcText.left = rcIcon.right + 4;
                                rcText.right = rcSubItem.right - 4;
                                DrawTextW(hdc, szText, -1, &rcText,
                                    DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
                            }
                            else {
                                RECT rcText = rcSubItem;
                                rcText.left += 4;
                                rcText.right -= 4;
                                UINT uFormat = DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX;
                                LVCOLUMN lvc = { LVCF_FMT };
                                if (ListView_GetColumn(hListView, iSubItem, &lvc)) {
                                    if (lvc.fmt & LVCFMT_RIGHT) {
                                        uFormat = DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX;
                                    }
                                    else if (lvc.fmt & LVCFMT_CENTER) {
                                        uFormat = DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX;
                                    }
                                }
                                DrawTextW(hdc, szText, -1, &rcText, uFormat);
                            }
                        }
                        if ((uItemState & CDIS_SELECTED) && (uItemState & CDIS_FOCUS)) {
                            DrawFocusRect(hdc, &rcItem);
                        }
                        return CDRF_SKIPDEFAULT;
                    }
                    case (CDDS_SUBITEM | CDDS_ITEMPREPAINT):
                        return CDRF_SKIPDEFAULT;
                    }
                }
                break;
                case LVN_GETDISPINFO:
                {
                    NMLVDISPINFO* plvdi = (NMLVDISPINFO*)lParam;
                    LVITEM* pItem = &(plvdi->item);
                    if (pItem->iItem >= (int)pData->fileData.size()) break;
                    FileInfo& info = pData->fileData[pItem->iItem];
                    const WIN32_FIND_DATAW& fd = info.findData;
                    if (pItem->mask & LVIF_TEXT) {
                        switch (pItem->iSubItem) {
                        case 0: StringCchCopy(pItem->pszText, pItem->cchTextMax, fd.cFileName); break;
                        case 1:
                            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
                                if (wcscmp(fd.cFileName, L"..") == 0) {
                                    StringCchCopy(pItem->pszText, pItem->cchTextMax, L"");
                                }
                                else {
                                    switch (info.sizeState) {
                                    case SizeCalculationState::InProgress:
                                    case SizeCalculationState::NotStarted:
                                        StringCchCopy(pItem->pszText, pItem->cchTextMax, L"計測中...");
                                        break;
                                    case SizeCalculationState::Completed:
                                        StrFormatByteSizeW(info.calculatedSize, pItem->pszText, pItem->cchTextMax);
                                        break;
                                    case SizeCalculationState::Error:
                                        StringCchCopy(pItem->pszText, pItem->cchTextMax, L"アクセス不可");
                                        break;
                                    }
                                }
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
                        case 4:
                        {
                            if (!info.sidString.empty()) {
                                EnterCriticalSection(&g_cacheLock);
                                auto it = g_sidNameCache.find(info.sidString);
                                if (it != g_sidNameCache.end()) {
                                    StringCchCopy(pItem->pszText, pItem->cchTextMax, it->second.c_str());
                                }
                                else {
                                    StringCchCopy(pItem->pszText, pItem->cchTextMax, L"取得中...");
                                }
                                LeaveCriticalSection(&g_cacheLock);
                            }
                            else {
                                StringCchCopy(pItem->pszText, pItem->cchTextMax, info.owner.c_str());
                            }
                            break;
                        }
                        }
                    }
                    if (pItem->mask & LVIF_IMAGE) {
                        if (info.iIcon == -1) {
                            WCHAR fullpath[MAX_PATH];
                            PathCombineW(fullpath, pData->currentPath, fd.cFileName);
                            SHFILEINFO sfi = { 0 };
                            SHGetFileInfoW(fullpath,
                                fd.dwFileAttributes,
                                &sfi,
                                sizeof(sfi),
                                SHGFI_USEFILEATTRIBUTES | SHGFI_SYSICONINDEX | SHGFI_SMALLICON);
                            info.iIcon = sfi.iIcon;
                            auto it = std::find_if(pData->allFileData.begin(), pData->allFileData.end(),
                                [&](const FileInfo& master_info) {
                                    return wcscmp(master_info.findData.cFileName, info.findData.cFileName) == 0;
                                });
                            if (it != pData->allFileData.end()) {
                                it->iIcon = sfi.iIcon;
                            }
                        }
                        pItem->iImage = info.iIcon;
                    }
                }
                break;
                case LVN_ENDLABELEDITW:
                {
                    NMLVDISPINFO* pdi = (NMLVDISPINFO*)lParam;
                    if (pdi->item.pszText != NULL && pdi->item.iItem != -1)
                    {
                        const FileInfo& info = pData->fileData[pdi->item.iItem];
                        WCHAR oldPath[MAX_PATH];
                        WCHAR newPath[MAX_PATH];
                        PathCombineW(oldPath, pData->currentPath, info.findData.cFileName);
                        PathCombineW(newPath, pData->currentPath, pdi->item.pszText);
                        if (MoveFileW(oldPath, newPath))
                        {
                            std::wstring newNameStr = pdi->item.pszText;
                            ListDirectory(hWnd, pData);
                            for (size_t i = 0; i < pData->fileData.size(); ++i) {
                                if (wcscmp(pData->fileData[i].findData.cFileName, newNameStr.c_str()) == 0) {
                                    ListView_SetItemState(pData->hList, (int)i, LVIS_FOCUSED | LVIS_SELECTED, LVIS_FOCUSED | LVIS_SELECTED);
                                    ListView_EnsureVisible(pData->hList, (int)i, FALSE);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            MessageBoxW(hWnd, L"名前の変更に失敗しました。\nファイルが開かれているか、同じ名前のファイルが既に存在する可能性があります。", L"エラー", MB_OK | MB_ICONERROR);
                            return FALSE;
                        }
                        return TRUE;
                    }
                    return FALSE;
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
                    auto sort_predicate = [&](const FileInfo& a, const FileInfo& b) {
                        return CompareFunction(a, b, pData);
                        };
                    std::sort(pData->allFileData.begin(), pData->allFileData.end(), sort_predicate);
                    std::sort(pData->fileData.begin(), pData->fileData.end(), sort_predicate);
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
                    case VK_F2:
                    {
                        int sel = ListView_GetNextItem(pData->hList, -1, LVNI_SELECTED | LVNI_FOCUSED);
                        if (sel != -1) {
                            if (wcscmp(pData->fileData[sel].findData.cFileName, L"..") != 0) {
                                ListView_EditLabel(pData->hList, sel);
                            }
                        }
                        break;
                    }
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
        if (g_hbrBg) DeleteObject(g_hbrBg);
        if (g_hbrHeaderBg) DeleteObject(g_hbrHeaderBg);
        for (const auto& tab : g_tabs) {
            if (tab->searchTimerId) KillTimer(hWnd, tab->searchTimerId);
        }
        DeleteObject(hFont);
        DeleteObject((HGDIOBJ)SendMessage(hAddButton, WM_GETFONT, 0, 0));
        DeleteCriticalSection(&g_cacheLock);
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
    pData->hList = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | LVS_REPORT | LVS_OWNERDATA | LVS_EDITLABELS,
        0, 0, 0, 0,
        hWnd, (HMENU)(LIST_ID_BASE + tabIndex), GetModuleHandle(NULL), NULL);
    if (!pData->hList) return;
    SetWindowTheme(pData->hList, g_isDarkMode ? L"DarkMode_Explorer" : L"Explorer", NULL);
    HWND hHeader = ListView_GetHeader(pData->hList);
    if (hHeader) {
        SetWindowSubclass(hHeader, HeaderSubclassProc, 0, 0);
    }
    SetWindowSubclass(pData->hList, ListSubclassProc, 0, (DWORD_PTR)pData.get());
    SendMessage(pData->hList, WM_SETFONT, (WPARAM)hFont, TRUE);
    ListView_SetExtendedListViewStyle(pData->hList, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_AUTOSIZECOLUMNS);
    ListView_SetImageList(pData->hList, hImgList, LVSIL_SMALL);
    ListView_SetBkColor(pData->hList, g_clrBg);
    ListView_SetTextColor(pData->hList, g_clrText);
    LVCOLUMNW lvc = { LVCF_TEXT | LVCF_WIDTH | LVCF_FMT };
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 250; lvc.pszText = (LPWSTR)L"名前"; ListView_InsertColumn(pData->hList, 0, &lvc);
    lvc.fmt = LVCFMT_RIGHT; lvc.cx = 100; lvc.pszText = (LPWSTR)L"サイズ"; ListView_InsertColumn(pData->hList, 1, &lvc);
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 150; lvc.pszText = (LPWSTR)L"更新日時"; ListView_InsertColumn(pData->hList, 2, &lvc);
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 150; lvc.pszText = (LPWSTR)L"作成日時"; ListView_InsertColumn(pData->hList, 3, &lvc);
    lvc.fmt = LVCFMT_LEFT; lvc.cx = 150; lvc.pszText = (LPWSTR)L"所有者"; ListView_InsertColumn(pData->hList, 4, &lvc);
    ListDirectory(hWnd, pData.get());
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
        ShowWindow(g_tabs[i]->hList, show ? SW_SHOW : SW_HIDE);
    }
    if (newIndex >= 0 && newIndex < (int)g_tabs.size()) {
        ExplorerTabData* pData = g_tabs[newIndex].get();
        pData->bProgrammaticPathChange = true;
        SetWindowTextW(hPathEdit, pData->currentPath);
        pData->bProgrammaticPathChange = false;
        SetFocus(pData->hList);
    }
    UpdateLayout(hWnd);
}
void UpdateLayout(HWND hWnd) {
    RECT rcClient;
    GetClientRect(hWnd, &rcClient);
    constexpr int TITLEBAR_HEIGHT = 32;
    constexpr int PATH_EDIT_WIDTH = 250;
    constexpr int PATH_EDIT_MARGIN = 4;
    constexpr int ADD_BUTTON_WIDTH = 22;
    constexpr int ADD_BUTTON_MARGIN = 4;
    RECT rcAdjust = { 0, 0, 0, 0 };
    TabCtrl_AdjustRect(hTab, TRUE, &rcAdjust);
    int tabHeight = -rcAdjust.top;
    int tabY = TITLEBAR_HEIGHT - tabHeight;
    int pathEditHeight = 24;
    int pathEditY = (TITLEBAR_HEIGHT - pathEditHeight) / 2;
    int addButtonY = (TITLEBAR_HEIGHT - ADD_BUTTON_WIDTH) / 2;
    int rightBoundary = g_rcMinButton.left;
    int pathEditX = rightBoundary - PATH_EDIT_MARGIN - PATH_EDIT_WIDTH;
    int addButtonX = pathEditX - ADD_BUTTON_MARGIN - ADD_BUTTON_WIDTH;
    int tabWidth = addButtonX;
    if (pathEditX < 0) pathEditX = 0;
    if (addButtonX < 0) addButtonX = 0;
    if (tabWidth < 0) tabWidth = 0;
    MoveWindow(hTab, 0, tabY, tabWidth, tabHeight, TRUE);
    MoveWindow(hAddButton, addButtonX, addButtonY, ADD_BUTTON_WIDTH, ADD_BUTTON_WIDTH, TRUE);
    MoveWindow(hPathEdit, pathEditX, pathEditY, PATH_EDIT_WIDTH, pathEditHeight, TRUE);
    SetWindowPos(hAddButton, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    SetWindowPos(hPathEdit, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
    int curSel = TabCtrl_GetCurSel(hTab);
    if (curSel != -1) {
        ExplorerTabData* pData = GetCurrentTabData();
        if (pData) {
            int contentTop = TITLEBAR_HEIGHT;
            MoveWindow(pData->hList, 0, contentTop, rcClient.right, rcClient.bottom - contentTop, TRUE);
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
    g_rcCloseButton = { rcClient.right - BTN_WIDTH, 0, rcClient.right, BTN_HEIGHT };
    g_rcMaxButton = { g_rcCloseButton.left - BTN_WIDTH, 0, g_rcCloseButton.left, BTN_HEIGHT };
    g_rcMinButton = { g_rcMaxButton.left - BTN_WIDTH, 0, g_rcMaxButton.left, BTN_HEIGHT };
}
void DrawCaptionButtons(HDC hdc, HWND hWnd) {
    COLORREF clrIcon, clrBg, clrHotBg, clrPressedBg, clrCloseHotBg, clrCloseIcon;
    if (g_isDarkMode) {
        clrIcon = RGB(255, 255, 255);
        clrBg = g_clrBg;
        clrHotBg = RGB(80, 80, 80);
        clrPressedBg = RGB(100, 100, 100);
    }
    else {
        clrIcon = RGB(0, 0, 0);
        clrBg = g_clrBg;
        clrHotBg = RGB(229, 243, 255);
        clrPressedBg = RGB(204, 232, 255);
    }
    clrCloseHotBg = RGB(232, 17, 35);
    clrCloseIcon = RGB(255, 255, 255);
    if (GetActiveWindow() != hWnd) {
        clrIcon = RGB(128, 128, 128);
    }
    HBRUSH hBrush = NULL;
    COLORREF minBg = clrBg;
    if (g_pressedButton == CaptionButtonState::Min && g_hoveredButton == CaptionButtonState::Min) minBg = clrPressedBg;
    else if (g_hoveredButton == CaptionButtonState::Min) minBg = clrHotBg;
    hBrush = CreateSolidBrush(minBg);
    FillRect(hdc, &g_rcMinButton, hBrush);
    DeleteObject(hBrush);
    HPEN hPen = CreatePen(PS_SOLID, 1, clrIcon);
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
    int minCenterX = g_rcMinButton.left + (g_rcMinButton.right - g_rcMinButton.left) / 2;
    int minCenterY = g_rcMinButton.top + (g_rcMinButton.bottom - g_rcMinButton.top) / 2;
    MoveToEx(hdc, minCenterX - 5, minCenterY, NULL);
    LineTo(hdc, minCenterX + 5, minCenterY);
    SelectObject(hdc, hOldPen);
    DeleteObject(hPen);
    COLORREF maxBg = clrBg;
    if (g_pressedButton == CaptionButtonState::Max && g_hoveredButton == CaptionButtonState::Max) maxBg = clrPressedBg;
    else if (g_hoveredButton == CaptionButtonState::Max) maxBg = clrHotBg;
    hBrush = CreateSolidBrush(maxBg);
    FillRect(hdc, &g_rcMaxButton, hBrush);
    DeleteObject(hBrush);
    hPen = CreatePen(PS_SOLID, 1, clrIcon);
    hOldPen = (HPEN)SelectObject(hdc, hPen);
    int maxCenterX = g_rcMaxButton.left + (g_rcMaxButton.right - g_rcMaxButton.left) / 2;
    int maxCenterY = g_rcMaxButton.top + (g_rcMaxButton.bottom - g_rcMaxButton.top) / 2;
    if (IsZoomed(hWnd)) {
        Rectangle(hdc, maxCenterX - 3, maxCenterY - 5, maxCenterX + 5, maxCenterY + 3);
        hBrush = CreateSolidBrush(maxBg);
        SelectObject(hdc, GetStockObject(NULL_PEN));
        SelectObject(hdc, hBrush);
        Rectangle(hdc, maxCenterX - 5, maxCenterY - 3, maxCenterX + 3, maxCenterY + 5);
        DeleteObject(hBrush);
        SelectObject(hdc, hPen);
        Rectangle(hdc, maxCenterX - 5, maxCenterY - 3, maxCenterX + 3, maxCenterY + 5);
    }
    else {
        Rectangle(hdc, maxCenterX - 5, maxCenterY - 5, maxCenterX + 5, maxCenterY + 5);
    }
    SelectObject(hdc, hOldPen);
    DeleteObject(hPen);
    COLORREF closeBg = clrBg;
    COLORREF currentCloseIcon = clrIcon;
    if (g_pressedButton == CaptionButtonState::Close && g_hoveredButton == CaptionButtonState::Close) {
        closeBg = clrCloseHotBg;
        currentCloseIcon = clrCloseIcon;
    }
    else if (g_hoveredButton == CaptionButtonState::Close) {
        closeBg = clrCloseHotBg;
        currentCloseIcon = clrCloseIcon;
    }
    hBrush = CreateSolidBrush(closeBg);
    FillRect(hdc, &g_rcCloseButton, hBrush);
    DeleteObject(hBrush);
    hPen = CreatePen(PS_SOLID, 2, currentCloseIcon);
    hOldPen = (HPEN)SelectObject(hdc, hPen);
    int closeCenterX = g_rcCloseButton.left + (g_rcCloseButton.right - g_rcCloseButton.left) / 2;
    int closeCenterY = g_rcCloseButton.top + (g_rcCloseButton.bottom - g_rcCloseButton.top) / 2;
    MoveToEx(hdc, closeCenterX - 5, closeCenterY - 5, NULL);
    LineTo(hdc, closeCenterX + 5, closeCenterY + 5);
    MoveToEx(hdc, closeCenterX - 5, closeCenterY + 5, NULL);
    LineTo(hdc, closeCenterX + 5, closeCenterY - 5);
    SelectObject(hdc, hOldPen);
    DeleteObject(hPen);
}
bool ShouldAppsUseDarkMode()
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        DWORD dwValue = 1;
        DWORD dwSize = sizeof(dwValue);
        if (RegQueryValueExW(hKey, L"AppsUseLightTheme", NULL, NULL, (LPBYTE)&dwValue, &dwSize) == ERROR_SUCCESS)
        {
            RegCloseKey(hKey);
            return dwValue == 0;
        }
        RegCloseKey(hKey);
    }
    return false;
}
void UpdateTheme(HWND hWnd)
{
    g_isDarkMode = ShouldAppsUseDarkMode();
    static fnAllowDarkModeForWindow pfnAllowDarkModeForWindow = nullptr;
    static bool api_loaded = false;
    if (!api_loaded) {
        HMODULE hUxtheme = LoadLibraryExW(L"uxtheme.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
        if (hUxtheme) {
            pfnAllowDarkModeForWindow = reinterpret_cast<fnAllowDarkModeForWindow>(GetProcAddress(hUxtheme, MAKEINTRESOURCEA(133)));
        }
        api_loaded = true;
    }
    BOOL useDarkMode = g_isDarkMode;
    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDarkMode, sizeof(useDarkMode));
    int backdropType = DWMSBT_MAINWINDOW;
    DwmSetWindowAttribute(hWnd, DWMWA_SYSTEM_BACKDROP_TYPE, &backdropType, sizeof(backdropType));
    if (g_isDarkMode) {
        g_clrBg = RGB(32, 32, 32);
        g_clrText = RGB(255, 255, 255);
        g_clrActiveTab = RGB(45, 45, 45);
        g_clrAccent = RGB(10, 132, 255);
        g_clrSeparator = RGB(60, 60, 60);
        g_clrCloseHoverBg = RGB(232, 17, 35);
        g_clrCloseHoverText = RGB(255, 255, 255);
        g_clrCloseText = RGB(200, 200, 200);
        g_clrSelection = RGB(38, 79, 120);
        g_clrSelectionText = RGB(255, 255, 255);
        g_clrHot = RGB(55, 55, 55);
        g_clrHeaderBg = RGB(45, 45, 45);
    }
    else {
        g_clrBg = RGB(240, 240, 240);
        g_clrText = RGB(0, 0, 0);
        g_clrActiveTab = RGB(255, 255, 255);
        g_clrAccent = RGB(0, 120, 215);
        g_clrSeparator = RGB(220, 220, 220);
        g_clrCloseHoverBg = RGB(232, 17, 35);
        g_clrCloseHoverText = RGB(255, 255, 255);
        g_clrCloseText = RGB(100, 100, 100);
        g_clrSelection = RGB(204, 232, 255);
        g_clrSelectionText = RGB(0, 0, 0);
        g_clrHot = RGB(229, 243, 255);
        g_clrHeaderBg = RGB(240, 240, 240);
    }
    if (g_hbrBg) DeleteObject(g_hbrBg);
    g_hbrBg = CreateSolidBrush(g_clrBg);
    if (g_hbrHeaderBg) DeleteObject(g_hbrHeaderBg);
    g_hbrHeaderBg = CreateSolidBrush(g_clrHeaderBg);
    if (pfnAllowDarkModeForWindow) {
        pfnAllowDarkModeForWindow(hPathEdit, g_isDarkMode);
    }
    SetWindowTheme(hPathEdit, L"", L"");
    SetWindowPos(hPathEdit, NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
    for (const auto& tab : g_tabs) {
        if (tab && tab->hList) {
            SetWindowTheme(tab->hList, g_isDarkMode ? L"DarkMode_Explorer" : L"Explorer", NULL);
            HWND hHeader = ListView_GetHeader(tab->hList);
            if (hHeader) SetWindowTheme(hHeader, L"", L"");
            ListView_SetBkColor(tab->hList, g_clrBg);
            ListView_SetTextColor(tab->hList, g_clrText);
            SetWindowPos(tab->hList, NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
        }
    }
    RedrawWindow(hWnd, NULL, NULL, RDW_ERASE | RDW_INVALIDATE | RDW_ALLCHILDREN);
}
LRESULT CALLBACK TabCtrlSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    switch (uMsg) {
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wParam;
        RECT rc;
        GetClientRect(hWnd, &rc);
        FillRect(hdc, &rc, g_hbrBg);
        return 1;
    }
    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, TabCtrlSubclassProc, uIdSubclass);
        break;
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPWSTR pCmdLine, int nCmdShow) {
    enum class AppMode { Default, AllowDark, ForceDark, ForceLight, Max };
    using fnSetPreferredAppMode = AppMode(WINAPI*)(AppMode appMode);
    HMODULE hUxtheme = LoadLibraryExW(L"uxtheme.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (hUxtheme) {
        auto pfnSetPreferredAppMode = reinterpret_cast<fnSetPreferredAppMode>(GetProcAddress(hUxtheme, MAKEINTRESOURCEA(135)));
        if (pfnSetPreferredAppMode) {
            pfnSetPreferredAppMode(AppMode::AllowDark);
        }
    }
    InitializeCriticalSection(&g_cacheLock);
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    INITCOMMONCONTROLSEX ic = { sizeof(ic), ICC_LISTVIEW_CLASSES | ICC_TAB_CLASSES };
    InitCommonControlsEx(&ic);
    MSG msg;
    WNDCLASSW wndclass = {
        CS_HREDRAW | CS_VREDRAW,
        WndProc, 0, 0, hInstance,
        LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)),
        LoadCursor(0, IDC_ARROW),
        0, 0, szClassName
    };
    RegisterClassW(&wndclass);
    HWND hWnd = CreateWindowEx(0, szClassName, L"fai",
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
int CalculateNameColumnWidth(ExplorerTabData* pData) {
    if (!pData || !pData->hList) return 250;
    HDC hdc = GetDC(pData->hList);
    if (!hdc) return 250;
    HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
    int maxWidth = 0;
    const int ICON_WIDTH = GetSystemMetrics(SM_CXSMICON);
    const int MARGIN = 8;
    const int ICON_TEXT_GAP = 4;
    SIZE headerSize;
    if (GetTextExtentPoint32W(hdc, L"名前", 2, &headerSize)) {
        maxWidth = headerSize.cx + MARGIN * 2;
    }
    for (const auto& fileInfo : pData->fileData) {
        SIZE textSize;
        int nameLen = wcslen(fileInfo.findData.cFileName);
        if (GetTextExtentPoint32W(hdc, fileInfo.findData.cFileName, nameLen, &textSize)) {
            int totalWidth = ICON_WIDTH + ICON_TEXT_GAP + textSize.cx + MARGIN * 2;
            if (totalWidth > maxWidth) {
                maxWidth = totalWidth;
            }
        }
    }
    SelectObject(hdc, hOldFont);
    ReleaseDC(pData->hList, hdc);
    const int MIN_WIDTH = 100;
    const int MAX_WIDTH = 500;
    if (maxWidth < MIN_WIDTH) maxWidth = MIN_WIDTH;
    if (maxWidth > MAX_WIDTH) maxWidth = MAX_WIDTH;
    return maxWidth;
}
LRESULT CALLBACK HeaderSubclassProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    switch (uMsg)
    {
    case WM_ERASEBKGND:
        return 1;
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        HFONT hOldFont = (HFONT)SelectObject(hdc, hFont);
        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        FillRect(hdc, &rcClient, g_hbrHeaderBg);
        int itemCount = Header_GetItemCount(hWnd);
        for (int i = 0; i < itemCount; ++i)
        {
            RECT rcItem;
            if (Header_GetItemRect(hWnd, i, &rcItem))
            {
                WCHAR szText[256];
                HDITEMW hdi = { HDI_TEXT | HDI_FORMAT, 0, szText, NULL, ARRAYSIZE(szText) };
                Header_GetItem(hWnd, i, &hdi);
                SetBkMode(hdc, TRANSPARENT);
                SetTextColor(hdc, g_clrText);
                RECT rcText = rcItem;
                rcText.left += 8;
                rcText.right -= 8;
                if (hdi.fmt & (HDF_SORTUP | HDF_SORTDOWN)) {
                    HPEN hPen = CreatePen(PS_SOLID, 1, g_clrText);
                    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                    HBRUSH hBrush = CreateSolidBrush(g_clrText);
                    HBRUSH hOldBrush = (HBRUSH)SelectObject(hdc, hBrush);
                    const int arrowWidth = 8;
                    const int arrowHeight = 4;
                    int midX = rcText.right - arrowWidth / 2 - 4;
                    int midY = rcItem.top + (rcItem.bottom - rcItem.top) / 2;
                    POINT points[3];
                    if (hdi.fmt & HDF_SORTUP) {
                        points[0] = { midX, midY - (arrowHeight / 2) };
                        points[1] = { midX - (arrowWidth / 2), midY + (arrowHeight / 2) };
                        points[2] = { midX + (arrowWidth / 2), midY + (arrowHeight / 2) };
                    }
                    else {
                        points[0] = { midX, midY + (arrowHeight / 2) };
                        points[1] = { midX - (arrowWidth / 2), midY - (arrowHeight / 2) };
                        points[2] = { midX + (arrowWidth / 2), midY - (arrowHeight / 2) };
                    }
                    Polygon(hdc, points, 3);
                    SelectObject(hdc, hOldPen);
                    DeleteObject(hPen);
                    SelectObject(hdc, hOldBrush);
                    DeleteObject(hBrush);
                    rcText.right -= (arrowWidth + 8);
                }
                DrawTextW(hdc, szText, -1, &rcText, DT_VCENTER | DT_SINGLELINE | DT_LEFT | DT_END_ELLIPSIS);
                HPEN hSepPen = CreatePen(PS_SOLID, 1, g_isDarkMode ? RGB(60, 60, 60) : RGB(220, 220, 220));
                HPEN hOldSepPen = (HPEN)SelectObject(hdc, hSepPen);
                MoveToEx(hdc, rcItem.right - 1, rcItem.top + 4, NULL);
                LineTo(hdc, rcItem.right - 1, rcItem.bottom - 4);
                SelectObject(hdc, hOldSepPen);
                DeleteObject(hSepPen);
            }
        }
        SelectObject(hdc, hOldFont);
        EndPaint(hWnd, &ps);
        return 0;
    }
    case WM_NCDESTROY:
        RemoveWindowSubclass(hWnd, HeaderSubclassProc, uIdSubclass);
        break;
    }
    return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}