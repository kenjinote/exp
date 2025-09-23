#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include <windows.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <shellapi.h>
#include <strsafe.h>
#include "resource.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")

WCHAR szClassName[] = L"exp";
static HWND hEdit, hList;
static HFONT hFont;
static HIMAGELIST hImgList;

void FileTimeToString(const FILETIME& ft, WCHAR* buf, size_t bufSize)
{
    SYSTEMTIME stUTC, stLocal;
    FileTimeToSystemTime(&ft, &stUTC);
    SystemTimeToTzSpecificLocalTime(NULL, &stUTC, &stLocal);
    StringCchPrintf(buf, bufSize, L"%04d/%02d/%02d %02d:%02d",
        stLocal.wYear, stLocal.wMonth, stLocal.wDay, stLocal.wHour, stLocal.wMinute);
}

void ListDirectory(HWND hList, LPCWSTR path)
{
    ListView_DeleteAllItems(hList);

    WIN32_FIND_DATA fd;
    WCHAR buf[MAX_PATH];
    wsprintf(buf, L"%s\\*.*", path);

    HANDLE hFind = FindFirstFile(buf, &fd);
    if (hFind == INVALID_HANDLE_VALUE) return;

    int i = 0;
    do {
        if (lstrcmp(fd.cFileName, L".") == 0) continue;

        WCHAR fullpath[MAX_PATH];
        wsprintf(fullpath, L"%s\\%s", path, fd.cFileName);

        SHFILEINFO sfi;
        int iIcon = 0;
        if (SHGetFileInfo(fullpath, 0, &sfi, sizeof(sfi),
            SHGFI_ICON | SHGFI_SMALLICON))
        {
            iIcon = ImageList_AddIcon(hImgList, sfi.hIcon);
            DestroyIcon(sfi.hIcon);
        }

        LVITEM item = { 0 };
        item.mask = LVIF_TEXT | LVIF_PARAM | LVIF_IMAGE;
        item.iItem = i++;
        item.pszText = fd.cFileName;
        item.iImage = iIcon;
        item.lParam = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ? 1 : 0;
        int index = ListView_InsertItem(hList, &item);

        // サイズ
        WCHAR sizebuf[64];
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            lstrcpy(sizebuf, L"<DIR>");
        }
        else {
            LARGE_INTEGER sz;
            sz.HighPart = fd.nFileSizeHigh;
            sz.LowPart = fd.nFileSizeLow;
            StringCchPrintf(sizebuf, 64, L"%lld", sz.QuadPart);
        }
        ListView_SetItemText(hList, index, 1, sizebuf);

        // 作成日時
        WCHAR timebuf[64];
        FileTimeToString(fd.ftCreationTime, timebuf, 64);
        ListView_SetItemText(hList, index, 2, timebuf);

        // 更新日時
        FileTimeToString(fd.ftLastWriteTime, timebuf, 64);
        ListView_SetItemText(hList, index, 3, timebuf);

    } while (FindNextFile(hFind, &fd));
    FindClose(hFind);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    static WCHAR currentPath[MAX_PATH];
    switch (msg)
    {
    case WM_CREATE:
    {
        hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0,
            WS_VISIBLE | WS_CHILD | ES_AUTOHSCROLL,
            0, 0, 0, 0, hWnd, (HMENU)1001,
            ((LPCREATESTRUCT)lParam)->hInstance, 0);

        hList = CreateWindow(WC_LISTVIEW, 0,
            WS_VISIBLE | WS_CHILD | LVS_REPORT | LVS_SINGLESEL,
            0, 0, 0, 0, hWnd, (HMENU)IDOK,
            ((LPCREATESTRUCT)lParam)->hInstance, 0);

        ListView_SetExtendedListViewStyle(hList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

        // カラム追加
        LVCOLUMN col = { LVCF_TEXT | LVCF_WIDTH };
        col.cx = 250; col.pszText = (LPWSTR)L"名前"; ListView_InsertColumn(hList, 0, &col);
        col.cx = 100; col.pszText = (LPWSTR)L"サイズ"; ListView_InsertColumn(hList, 1, &col);
        col.cx = 150; col.pszText = (LPWSTR)L"作成日時"; ListView_InsertColumn(hList, 2, &col);
        col.cx = 150; col.pszText = (LPWSTR)L"更新日時"; ListView_InsertColumn(hList, 3, &col);

        hFont = CreateFontW(18, 0, 0, 0, FW_NORMAL, 0, 0, 0,
            SHIFTJIS_CHARSET, 0, 0, 0, 0, L"Segoe UI");
        SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, 0);
        SendMessage(hList, WM_SETFONT, (WPARAM)hFont, 0);

        // アイコンイメージリスト
        hImgList = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 256);
        ListView_SetImageList(hList, hImgList, LVSIL_SMALL);

        GetCurrentDirectory(MAX_PATH, currentPath);
        SetWindowText(hEdit, currentPath);
        ListDirectory(hList, currentPath);
    }
    break;

    case WM_SIZE:
        MoveWindow(hEdit, 0, 0, LOWORD(lParam), 30, TRUE);
        MoveWindow(hList, 0, 30, LOWORD(lParam), HIWORD(lParam) - 30, TRUE);
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == 1001 && HIWORD(wParam) == EN_UPDATE)
        {
            GetWindowText(hEdit, currentPath, MAX_PATH);
            // PathCanonicalizeWを使って..を解決する
			WCHAR resolvedPath[MAX_PATH];
			if (PathCanonicalizeW(resolvedPath, currentPath)) {
				lstrcpy(currentPath, resolvedPath);
			}
            if (PathIsDirectory(currentPath)) {
                ListDirectory(hList, currentPath);
            }
        }
        break;

    case WM_NOTIFY:
        if (((LPNMHDR)lParam)->hwndFrom == hList &&
            ((LPNMHDR)lParam)->code == NM_DBLCLK)
        {
            int sel = ListView_GetNextItem(hList, -1, LVNI_SELECTED);
            if (sel >= 0) {
                WCHAR buf[MAX_PATH];
                ListView_GetItemText(hList, sel, 0, buf, MAX_PATH);

                WCHAR fullpath[MAX_PATH];
                wsprintf(fullpath, L"%s\\%s", currentPath, buf);

                DWORD attr = GetFileAttributes(fullpath);
                if (attr & FILE_ATTRIBUTE_DIRECTORY) {
                    WCHAR resolvedPath[MAX_PATH];
                    if (PathCanonicalizeW(resolvedPath, fullpath)) {
                        lstrcpy(currentPath, resolvedPath);
					}
					else {
						lstrcpy(currentPath, fullpath);
					}
                    SetWindowText(hEdit, currentPath);
                    ListDirectory(hList, currentPath);
                }
                else {
                    ShellExecute(NULL, L"open", fullpath, NULL, NULL, SW_SHOWNORMAL);
                }
            }
        }
        break;

    case WM_DESTROY:
        DeleteObject(hFont);
        if (hImgList) ImageList_Destroy(hImgList);
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPWSTR pCmdLine, int nCmdShow)
{
    INITCOMMONCONTROLSEX ic = { sizeof(ic), ICC_LISTVIEW_CLASSES };
    InitCommonControlsEx(&ic);

    MSG msg;
    WNDCLASS wndclass = {
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
    RegisterClass(&wndclass);

    HWND hWnd = CreateWindow(
        szClassName, L"超高速シンプルファイラ",
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT, 0, 900, 600,
        0, 0, hInstance, 0);

    ShowWindow(hWnd, SW_SHOWDEFAULT);
    UpdateWindow(hWnd);

    while (GetMessage(&msg, 0, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}
