#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <wchar.h>
#include <assert.h>
#include "resource.h"

#define WM_CHANGE_NOTIFY (WM_USER + 125)
#define TIMER_ID_NOTIFY 999

typedef struct DATA
{
    HINSTANCE hInst;
    HWND hwnd;
    HWND hwndToolbar;
    HWND hwndTreeView;
    HIMAGELIST hToolbarImageList;
    HIMAGELIST hTreeViewImageList;
    INT iWebPageImage;
    UINT uNotifyReg;
} DATA, *PDATA;

LPCITEMIDLIST GetPidlFromTreeViewItem(PDATA pData, HTREEITEM hItem)
{
    TV_ITEMW item = { TVIF_PARAM };
    item.hItem = hItem;
    TreeView_GetItem(pData->hwndTreeView, &item);
    return (LPCITEMIDLIST)item.lParam;
}

void PopulateTreeView(PDATA pData, HTREEITEM hParent, LPCWSTR pszParent)
{
    WCHAR szPath[MAX_PATH];
    lstrcpynW(szPath, pszParent, MAX_PATH);
    PathAppendW(szPath, L"*");

    WIN32_FIND_DATAW find;
    HANDLE hFind = FindFirstFileW(szPath, &find);
    if (hFind == INVALID_HANDLE_VALUE)
        return;

    do
    {
        if (find.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN)
            continue;
        if (lstrcmpW(find.cFileName, L".") == 0 || lstrcmpW(find.cFileName, L"..") == 0)
            continue;

        PathRemoveFileSpecW(szPath);
        PathAppendW(szPath, find.cFileName);

        SHFILEINFOW file_info = { NULL };
        SHGetFileInfoW(szPath, find.dwFileAttributes, &file_info, sizeof(file_info),
                       SHGFI_USEFILEATTRIBUTES | SHGFI_DISPLAYNAME | SHGFI_ICON | SHGFI_SMALLICON);

        TV_ITEMW item = { TVIF_TEXT | TVIF_IMAGE | TVIF_SELECTEDIMAGE | TVIF_PARAM };
        item.pszText = file_info.szDisplayName;
        item.cchTextMax = _countof(file_info.szDisplayName);
        item.lParam = (LPARAM)ILCreateFromPath(szPath);
        if (lstrcmpiW(PathFindExtensionW(szPath), L".url") == 0)
        {
            item.iImage = item.iSelectedImage = pData->iWebPageImage;
        }
        else
        {
            item.iImage = ImageList_AddIcon(pData->hTreeViewImageList, file_info.hIcon);
            item.iSelectedImage = item.iImage;
        }
        DestroyIcon(file_info.hIcon);

        TV_INSERTSTRUCTW insert = { hParent, TVI_LAST };
        insert.item = item;
        HTREEITEM hItem = TreeView_InsertItem(pData->hwndTreeView, &insert);

        if (find.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            PopulateTreeView(pData, hItem, szPath);
        }
    } while (FindNextFileW(hFind, &find));

    FindClose(hFind);
}

HTREEITEM GetTreeViewItemByPath(PDATA pData, HTREEITEM hItem, LPCWSTR pszSelect)
{
    if (!hItem)
        return NULL;

    for (; hItem; hItem = TreeView_GetNextSibling(pData->hwndTreeView, hItem))
    {
        WCHAR szPath[MAX_PATH];
        LPCITEMIDLIST pidl = GetPidlFromTreeViewItem(pData, hItem);
        if (SHGetPathFromIDListW(pidl, szPath) && lstrcmpiW(szPath, pszSelect) == 0)
            return hItem;

        HTREEITEM hChild = TreeView_GetChild(pData->hwndTreeView, hItem);
        if (!hChild)
            continue;

        HTREEITEM hSelected = GetTreeViewItemByPath(pData, hChild, pszSelect);
        if (hSelected)
            return hSelected;
    }

    return NULL;
}

void AddWebPageImage(PDATA pData, HIMAGELIST hImageList)
{
    SHFILEINFOW file_info = { NULL };
    SHGetFileInfoW(L".html", FILE_ATTRIBUTE_NORMAL, &file_info, sizeof(file_info),
                   SHGFI_USEFILEATTRIBUTES | SHGFI_ICON | SHGFI_SMALLICON);
    pData->iWebPageImage = ImageList_AddIcon(hImageList, file_info.hIcon);
    DestroyIcon(file_info.hIcon);
}

void RefreshTreeView(PDATA pData)
{
    WCHAR szFavorites[MAX_PATH];
    SHGetSpecialFolderPathW(pData->hwnd, szFavorites, CSIDL_FAVORITES, TRUE);

    WCHAR szPath[MAX_PATH] = { 0 };
    HTREEITEM hSelected = TreeView_GetSelection(pData->hwndTreeView);
    BOOL bExpanded = (TreeView_GetItemState(pData->hwndTreeView, hSelected, TVIS_EXPANDED) & TVIS_EXPANDED);
    if (hSelected)
    {
        LPCITEMIDLIST pidl = GetPidlFromTreeViewItem(pData, hSelected);
        if (!SHGetPathFromIDListW(pidl, szPath))
            szPath[0] = 0;
    }

    ImageList_Destroy(pData->hTreeViewImageList);
    pData->hTreeViewImageList = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 64, 0);
    TreeView_SetImageList(pData->hwndTreeView, pData->hTreeViewImageList, TVSIL_NORMAL);
    AddWebPageImage(pData, pData->hTreeViewImageList);

    SetWindowRedraw(pData->hwndTreeView, FALSE);
    {
        TreeView_DeleteAllItems(pData->hwndTreeView);

        PopulateTreeView(pData, TVI_ROOT, szFavorites);

        if (szPath[0])
        {
            HTREEITEM hRoot = TreeView_GetRoot(pData->hwndTreeView);
            HTREEITEM hSelect = GetTreeViewItemByPath(pData, hRoot, szPath);
            if (hSelect)
            {
                TreeView_SelectItem(pData->hwndTreeView, hSelect);
                if (bExpanded)
                    TreeView_Expand(pData->hwndTreeView, hSelect, TVE_EXPAND);
            }
        }
    }
    SetWindowRedraw(pData->hwndTreeView, TRUE);
}

BOOL InitFavoritesTreeView(PDATA pData)
{
    pData->hTreeViewImageList = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 64, 0);
    if (!pData->hTreeViewImageList)
        return FALSE;

    TreeView_SetImageList(pData->hwndTreeView, pData->hTreeViewImageList, TVSIL_NORMAL);
    AddWebPageImage(pData, pData->hTreeViewImageList);

    WCHAR szFavorites[MAX_PATH];
    SHGetSpecialFolderPathW(pData->hwnd, szFavorites, CSIDL_FAVORITES, TRUE);

    SetWindowRedraw(pData->hwndTreeView, FALSE);
    PopulateTreeView(pData, TVI_ROOT, szFavorites);
    SetWindowRedraw(pData->hwndTreeView, TRUE);
    return TRUE;
}

BOOL OnCreate(HWND hwnd, LPCREATESTRUCT lpCreateStruct)
{
    PDATA pData = (PDATA)lpCreateStruct->lpCreateParams;
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)pData);
    pData->hwnd = hwnd;

    HBITMAP hbmToolbar = LoadBitmapW(pData->hInst, MAKEINTRESOURCEW(214));
    assert(hbmToolbar);

    pData->hToolbarImageList = ImageList_Create(24, 24, ILC_COLOR32, 0, 8);
    if (!pData->hToolbarImageList)
    {
        assert(0);
        return FALSE;
    }
    ImageList_Add(pData->hToolbarImageList, hbmToolbar, NULL);
    DeleteObject(hbmToolbar);

    DWORD style, exstyle;

    style = WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT | TBSTYLE_LIST | CCS_NODIVIDER | TBSTYLE_WRAPABLE;
    exstyle = 0;
    pData->hwndToolbar = CreateWindowExW(exstyle, TOOLBARCLASSNAMEW, NULL, style,
                                         0, 0, 0, 0,
                                         hwnd, (HMENU)(ULONG_PTR)IDW_TOOLBAR,
                                         pData->hInst, NULL);
    if (!pData->hwndToolbar)
    {
        assert(0);
        ImageList_Destroy(pData->hToolbarImageList);
        return FALSE;
    }
    SendMessageW(pData->hwndToolbar, TB_BUTTONSTRUCTSIZE, (WPARAM)sizeof(TBBUTTON), 0);
    SendMessageW(pData->hwndToolbar, TB_SETIMAGELIST, 0, (LPARAM)pData->hToolbarImageList);
    SendMessageW(pData->hwndToolbar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_MIXEDBUTTONS);

    WCHAR szzAdd[MAX_PATH], szzOrganize[MAX_PATH];
    ZeroMemory(szzAdd, sizeof(szzAdd));
    ZeroMemory(szzOrganize, sizeof(szzOrganize));
    LoadStringW(pData->hInst, IDS_ADD, szzAdd, _countof(szzAdd));
    LoadStringW(pData->hInst, IDS_ORGANIZE, szzOrganize, _countof(szzOrganize));

    TBBUTTON tbb[2] = { { 0 } };
    INT iButton = 0;
    tbb[iButton].iBitmap = 3;
    tbb[iButton].idCommand = ID_ADD;
    tbb[iButton].fsState = TBSTATE_ENABLED;
    tbb[iButton].fsStyle = BTNS_BUTTON | BTNS_AUTOSIZE | BTNS_SHOWTEXT;
    tbb[iButton].iString = (INT)SendMessageW(pData->hwndToolbar, TB_ADDSTRING, 0, (LPARAM)szzAdd);
    ++iButton;
    tbb[iButton].iBitmap = 42;
    tbb[iButton].idCommand = ID_ORGANIZE;
    tbb[iButton].fsState = TBSTATE_ENABLED;
    tbb[iButton].fsStyle = BTNS_BUTTON | BTNS_AUTOSIZE | BTNS_SHOWTEXT;
    tbb[iButton].iString = (INT)SendMessageW(pData->hwndToolbar, TB_ADDSTRING, 0, (LPARAM)szzOrganize);
    ++iButton;
    assert(iButton == _countof(tbb));
    LRESULT ret = SendMessageW(pData->hwndToolbar, TB_ADDBUTTONS, iButton, (LPARAM)&tbb);
    assert(ret);

    style = TVS_NOHSCROLL | TVS_NONEVENHEIGHT | TVS_FULLROWSELECT | TVS_INFOTIP |
            TVS_SINGLEEXPAND | TVS_TRACKSELECT | TVS_SHOWSELALWAYS | TVS_EDITLABELS |
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_TABSTOP;
    exstyle = WS_EX_CLIENTEDGE;
    pData->hwndTreeView = CreateWindowExW(exstyle, WC_TREEVIEWW, NULL, style,
                                          0, 0, 0, 0,
                                          hwnd, (HMENU)(ULONG_PTR)IDW_TREEVIEW,
                                          pData->hInst, NULL);
    if (!pData->hwndTreeView)
    {
        assert(0);
        DestroyWindow(pData->hwndToolbar);
        ImageList_Destroy(pData->hToolbarImageList);
        return FALSE;
    }

    InitFavoritesTreeView(pData);

    LPITEMIDLIST pidlFavs;
    SHGetSpecialFolderLocation(hwnd, CSIDL_FAVORITES, &pidlFavs);
    SHChangeNotifyEntry entry = { pidlFavs, TRUE }; 
    pData->uNotifyReg = SHChangeNotifyRegister(hwnd, 
                                               SHCNRF_NewDelivery | SHCNRF_ShellLevel,
                                               SHCNE_DISKEVENTS,
                                               WM_CHANGE_NOTIFY,
                                               1, &entry);
    return TRUE;
}

PDATA GetData(HWND hwnd)
{
    return (PDATA)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
}

void OnOrganize(HWND hwnd)
{
    LPITEMIDLIST pidlFavs;
    HRESULT hr = SHGetSpecialFolderLocation(hwnd, CSIDL_FAVORITES, &pidlFavs);
    if (FAILED(hr))
        return;

    SHELLEXECUTEINFO info = { sizeof(info) };
    info.fMask = SEE_MASK_IDLIST;
    info.hwnd = hwnd;
    info.lpIDList = pidlFavs;
    info.nShow = SW_SHOWNORMAL;
    ShellExecuteExW(&info);
    ILFree(pidlFavs);
}

BOOL OnClick(HWND hwnd, HTREEITEM hItem)
{
    PDATA pData = GetData(hwnd);
    if (!pData || !hItem)
        return FALSE;

    WCHAR szPath[MAX_PATH];
    LPCITEMIDLIST pidl = GetPidlFromTreeViewItem(pData, hItem);
    if (!pidl || !SHGetPathFromIDListW(pidl, szPath) || !PathFileExistsW(szPath))
        return FALSE;

    if (PathIsDirectoryW(szPath))
        return FALSE;

    TreeView_SelectItem(pData->hwndTreeView, hItem);
    ShellExecuteW(hwnd, NULL, szPath, NULL, NULL, SW_SHOWNORMAL);
    return TRUE;
}

BOOL OnRightClick(HWND hwnd, HTREEITEM hItem)
{
    PDATA pData = GetData(hwnd);
    TreeView_SelectItem(pData->hwndTreeView, hItem);
    // FIXME
    return TRUE;
}

void OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
    PDATA pData = GetData(hwnd);
    switch (id)
    {
        case ID_ADD:
            MessageBoxW(hwnd, L"ID_ADD", L"ID_ADD", 0);
            break;
        case ID_ORGANIZE:
            OnOrganize(hwnd);
            break;
    }
}

void OnSize(HWND hwnd, UINT state, int cx, int cy)
{
    PDATA pData = GetData(hwnd);
    if (!pData)
        return;

    RECT rcClient;
    GetClientRect(hwnd, &rcClient);
    cx = rcClient.right;
    cy = rcClient.bottom;

    RECT rcTB;
    SendMessageW(pData->hwndToolbar, TB_AUTOSIZE, 0, 0);
    GetWindowRect(pData->hwndToolbar, &rcTB);

    INT cyTB = rcTB.bottom - rcTB.top;
    MoveWindow(pData->hwndTreeView, 0, cyTB, cx, cy - cyTB, TRUE);
}

LRESULT OnNotify(HWND hwnd, int idFrom, LPNMHDR pnmhdr)
{
    PDATA pData = GetData(hwnd);

    switch (pnmhdr->code)
    {
        case NM_CLICK:
        {
            TV_HITTESTINFO hit;
            GetCursorPos(&hit.pt);
            ScreenToClient(pData->hwndTreeView, &hit.pt);
            HTREEITEM hItem = TreeView_HitTest(pData->hwndTreeView, &hit);
            if (hItem && !(hit.flags & TVHT_NOWHERE))
                return OnClick(hwnd, hItem);
            break;
        }
        case NM_RCLICK: // Right-Click
        {
            TV_HITTESTINFO hit;
            GetCursorPos(&hit.pt);
            ScreenToClient(pData->hwndTreeView, &hit.pt);
            HTREEITEM hItem = TreeView_HitTest(pData->hwndTreeView, &hit);
            if (hItem && !(hit.flags & TVHT_NOWHERE))
                return OnRightClick(hwnd, hItem);
            break;
        }
        case NM_RETURN: // Enter key
        {
            if (pnmhdr->hwndFrom != pData->hwndTreeView)
                break;

            HTREEITEM hItem = TreeView_GetSelection(pData->hwndTreeView);
            return OnClick(hwnd, hItem);
        }
        case TVN_DELETEITEM:
        {
            NM_TREEVIEW *pTreeView = (NM_TREEVIEW *)pnmhdr;
            CoTaskMemFree((LPITEMIDLIST)pTreeView->itemOld.lParam);
            break;
        }
        case TVN_KEYDOWN:
        {
            TV_KEYDOWN *pKeyDown = (TV_KEYDOWN *)pnmhdr;
            if (pKeyDown->wVKey == L'R' && GetKeyState(VK_CONTROL) < 0) // Ctrl+R: Refresh
            {
                RefreshTreeView(pData);
                break;
            }
            if (pKeyDown->wVKey == VK_F5) // F5: Refresh
            {
                RefreshTreeView(pData);
                break;
            }
            if (pKeyDown->wVKey == VK_F2) // F2: Rename
            {
                HTREEITEM hItem = TreeView_GetSelection(pData->hwndTreeView);
                TreeView_EditLabel(pData->hwndTreeView, hItem);
                break;
            }
            if (pKeyDown->wVKey == VK_DELETE) // Delete key
            {
                HTREEITEM hItem = TreeView_GetSelection(pData->hwndTreeView);
                LPCITEMIDLIST pidl = GetPidlFromTreeViewItem(pData, hItem);
                WCHAR szzPath[MAX_PATH] = { 0 };
                SHGetPathFromIDListW(pidl, szzPath);
                SHFILEOPSTRUCTW file_op = { hwnd, FO_DELETE, szzPath };
                file_op.fFlags = FOF_ALLOWUNDO;
                SHFileOperationW(&file_op);
                break;
            }
            break;
        }
        case TVN_ENDLABELEDIT:
        {
            TV_DISPINFO *pDispInfo = (TV_DISPINFO *)pnmhdr;
            if (pDispInfo->item.pszText)
            {
                HTREEITEM hItem = TreeView_GetSelection(pData->hwndTreeView);
                LPCITEMIDLIST pidl = GetPidlFromTreeViewItem(pData, hItem);
                WCHAR szzPath[MAX_PATH] = { 0 };
                SHGetPathFromIDListW(pidl, szzPath);
                WCHAR szDest[MAX_PATH];
                lstrcpynW(szDest, szzPath, _countof(szDest));
                PathRemoveFileSpecW(szDest);
                PathAppendW(szDest, pDispInfo->item.pszText);
                SHFILEOPSTRUCTW file_op = { hwnd, FO_RENAME, szzPath, szDest };
                file_op.fFlags = FOF_ALLOWUNDO;
                if (SHFileOperationW(&file_op) == 0)
                    return TRUE;
            }
            break;
        }
    }
    return 0;
}

void OnDestroy(HWND hwnd)
{
    PDATA pData = GetData(hwnd);
    if (!pData)
        return;

    SHChangeNotifyDeregister(pData->uNotifyReg);
    ImageList_Destroy(pData->hToolbarImageList);
    ImageList_Destroy(pData->hTreeViewImageList);
    DestroyWindow(pData->hwndToolbar);
    DestroyWindow(pData->hwndTreeView);
    PostQuitMessage(0);
}

void OnChangeNotify(HWND hwnd, WPARAM wParam, LPARAM lParam)
{
    // Add delay for performance
    KillTimer(hwnd, TIMER_ID_NOTIFY);
    SetTimer(hwnd, TIMER_ID_NOTIFY, 1000, NULL);
}

void OnTimer(HWND hwnd, UINT id)
{
    if (id != TIMER_ID_NOTIFY)
        return;

    KillTimer(hwnd, id);
    PDATA pData = GetData(hwnd);
    RefreshTreeView(pData);
}

void OnActivate(HWND hwnd, UINT state, HWND hwndActDeact, BOOL fMinimized)
{
    PDATA pData = GetData(hwnd);
    SetFocus(pData->hwndTreeView);
}

void OnDropFiles(HWND hwnd, HDROP hdrop)
{
    PDATA pData = GetData(hwnd);

    WCHAR szzPath[MAX_PATH] = { 0 }, szDest[MAX_PATH];
    TV_HITTESTINFO hit;
    if (!DragQueryFileW(hdrop, 0, szzPath, _countof(szzPath)) || !DragQueryPoint(hdrop, &hit.pt))
        return;

    DragFinish(hdrop);

    MapWindowPoints(hwnd, pData->hwndTreeView, &hit.pt, 1);
    HTREEITEM hItem = TreeView_HitTest(pData->hwndTreeView, &hit);
    if (!hItem) // No item hit
    {
        SHGetSpecialFolderPathW(pData->hwnd, szDest, CSIDL_FAVORITES, TRUE);
    }
    else
    {
        LPCITEMIDLIST pidl = GetPidlFromTreeViewItem(pData, hItem);
        if (!SHGetPathFromIDListW(pidl, szDest))
            return;
    }
    PathAppendW(szDest, PathFindFileNameW(szzPath));

    SHFILEOPSTRUCTW file_op = { hwnd, FO_COPY, szzPath, szDest };
    file_op.fFlags = FOF_ALLOWUNDO;
    SHFileOperationW(&file_op);
}

LRESULT CALLBACK
WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        HANDLE_MSG(hwnd, WM_CREATE, OnCreate);
        HANDLE_MSG(hwnd, WM_DESTROY, OnDestroy);
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
        HANDLE_MSG(hwnd, WM_NOTIFY, OnNotify);
        HANDLE_MSG(hwnd, WM_SIZE, OnSize);
        HANDLE_MSG(hwnd, WM_ACTIVATE, OnActivate);
        HANDLE_MSG(hwnd, WM_TIMER, OnTimer);
        HANDLE_MSG(hwnd, WM_DROPFILES, OnDropFiles);
    case WM_CHANGE_NOTIFY:
        OnChangeNotify(hwnd, wParam, lParam);
        break;
    default:
        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

INT WINAPI
WinMain(HINSTANCE   hInstance,
         HINSTANCE  hPrevInstance,
         LPSTR      lpCmdLine,
         INT        nCmdShow)
{
    InitCommonControls();
    DATA data = { hInstance };

    WNDCLASSW wc = { CS_DBLCLKS | CS_HREDRAW | CS_VREDRAW };
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIconW(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
    wc.lpszClassName = L"katahiromz's Favorites mock-up";
    if (!RegisterClassW(&wc))
    {
        assert(0);
        return 1;
    }

    DWORD style = WS_OVERLAPPEDWINDOW, exstyle = WS_EX_TOOLWINDOW | WS_EX_ACCEPTFILES;
    HWND hwnd = CreateWindowExW(exstyle, wc.lpszClassName, L"Favorites", style,
        CW_USEDEFAULT, CW_USEDEFAULT, 250, 400,
        NULL, NULL, hInstance, &data);
    if (!hwnd)
    {
        assert(0);
        return 1;
    }
    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return (INT)msg.wParam;
}
