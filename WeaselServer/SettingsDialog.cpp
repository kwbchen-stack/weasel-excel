#include "stdafx.h"
#include "SettingsDialog.h"
#include <commdlg.h>
#include <atomic>
#include <thread>
#include <mutex>

#pragma comment(lib, "comdlg32.lib")

namespace excel_edition {

namespace {

constexpr wchar_t kClassName[] = L"WeaselExcelEditionSettings";

enum ControlId {
  IDC_PINYIN = 1001,
  IDC_SINGLE_CHAR,
  IDC_HORIZONTAL,
  IDC_COMPLETION,
  IDC_FILE,
  IDC_BROWSE,
  IDC_SHEET,
  IDC_INFO,
  IDC_STATS,
  IDC_APPLY,
};

std::mutex g_mutex;
HWND g_window = NULL;       // the open window, if any
bool g_starting = false;    // a window thread is being started

struct SettingsWindow {
  SettingsBackend backend;
  SettingsState state;  // what is currently applied
  HWND hwnd = NULL;
  HFONT font = NULL;
  int dpi = 96;

  int Px(int v) const { return MulDiv(v, dpi, 96); }

  HWND Add(const wchar_t* cls,
           const wchar_t* text,
           DWORD style,
           int x,
           int y,
           int w,
           int h,
           int id,
           DWORD ex_style = 0) {
    HWND c = CreateWindowExW(ex_style, cls, text, WS_CHILD | WS_VISIBLE | style,
                             Px(x), Px(y), Px(w), Px(h), hwnd,
                             (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL);
    if (c && font)
      SendMessageW(c, WM_SETFONT, (WPARAM)font, TRUE);
    return c;
  }

  HWND Item(int id) const { return GetDlgItem(hwnd, id); }

  void SetCheck(int id, bool on) {
    SendMessageW(Item(id), BM_SETCHECK, on ? BST_CHECKED : BST_UNCHECKED, 0);
  }
  bool GetCheck(int id) const {
    return SendMessageW(Item(id), BM_GETCHECK, 0, 0) == BST_CHECKED;
  }
  std::wstring GetText(int id) const {
    int len = GetWindowTextLengthW(Item(id));
    std::wstring s(len + 1, L'\0');
    GetWindowTextW(Item(id), &s[0], len + 1);
    s.resize(len);
    return s;
  }
  static std::wstring Trim(const std::wstring& s) {
    size_t b = s.find_first_not_of(L" \t\r\n\"");
    if (b == std::wstring::npos)
      return std::wstring();
    size_t e = s.find_last_not_of(L" \t\r\n\"");
    return s.substr(b, e - b + 1);
  }

  void ShowState(const SettingsState& s) {
    SetCheck(IDC_PINYIN, s.pinyin);
    SetCheck(IDC_SINGLE_CHAR, s.single_char);
    SetCheck(IDC_HORIZONTAL, s.horizontal);
    SetCheck(IDC_COMPLETION, s.completion);
    SetWindowTextW(Item(IDC_FILE), s.excel_file.c_str());
    SetWindowTextW(Item(IDC_SHEET), s.sheet.c_str());
  }

  SettingsState ReadState() const {
    SettingsState s;
    s.pinyin = GetCheck(IDC_PINYIN);
    s.single_char = GetCheck(IDC_SINGLE_CHAR);
    s.horizontal = GetCheck(IDC_HORIZONTAL);
    s.completion = GetCheck(IDC_COMPLETION);
    s.excel_file = Trim(GetText(IDC_FILE));
    s.sheet = Trim(GetText(IDC_SHEET));
    return s;
  }

  void Create() {
    HDC dc = GetDC(NULL);
    dpi = GetDeviceCaps(dc, LOGPIXELSY);
    ReleaseDC(NULL, dc);
    if (dpi <= 0)
      dpi = 96;

    NONCLIENTMETRICSW ncm = {};
    ncm.cbSize = sizeof(ncm);
    if (SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0))
      font = CreateFontIndirectW(&ncm.lfMessageFont);

    const DWORD tab = WS_TABSTOP;
    // 输入方案
    Add(L"BUTTON", L"输入方案", BS_GROUPBOX, 12, 8, 496, 50, 0);
    Add(L"BUTTON", L"拼音模式（不勾选 = 小兵五笔86）", BS_AUTOCHECKBOX | tab, 28,
        28, 460, 20, IDC_PINYIN);
    // 候选
    Add(L"BUTTON", L"候选", BS_GROUPBOX, 12, 64, 496, 104, 0);
    Add(L"BUTTON", L"单字模式（只显示单字，不显示词组）",
        BS_AUTOCHECKBOX | tab, 28, 84, 460, 20, IDC_SINGLE_CHAR);
    Add(L"BUTTON", L"候选横排", BS_AUTOCHECKBOX | tab, 28, 108, 460, 20,
        IDC_HORIZONTAL);
    Add(L"BUTTON", L"逐码提示（编码没打完时，也显示以已打编码开头的字词）",
        BS_AUTOCHECKBOX | tab, 28, 132, 460, 20, IDC_COMPLETION);
    // Excel 码表
    Add(L"BUTTON", L"Excel 码表", BS_GROUPBOX, 12, 174, 496, 128, 0);
    Add(L"STATIC", L"Excel 文件（.xlsx）：", SS_LEFT, 28, 194, 300, 18, 0);
    Add(L"EDIT", L"", ES_AUTOHSCROLL | tab, 28, 214, 386, 24, IDC_FILE,
        WS_EX_CLIENTEDGE);
    Add(L"BUTTON", L"浏览…", BS_PUSHBUTTON | tab, 422, 213, 76, 26, IDC_BROWSE);
    Add(L"STATIC", L"工作表名称：", SS_LEFT, 28, 250, 90, 18, 0);
    Add(L"EDIT", L"", ES_AUTOHSCROLL | tab, 120, 247, 230, 24, IDC_SHEET,
        WS_EX_CLIENTEDGE);
    Add(L"STATIC", L"", SS_LEFT, 28, 278, 470, 18, IDC_INFO);
    // buttons
    Add(L"BUTTON", L"查看打字统计…", BS_PUSHBUTTON | tab, 12, 316, 130, 28,
        IDC_STATS);
    Add(L"BUTTON", L"确定", BS_DEFPUSHBUTTON | tab, 262, 316, 78, 28, IDOK);
    Add(L"BUTTON", L"取消", BS_PUSHBUTTON | tab, 346, 316, 78, 28, IDCANCEL);
    Add(L"BUTTON", L"应用", BS_PUSHBUTTON | tab, 430, 316, 78, 28, IDC_APPLY);

    state = backend.load ? backend.load() : SettingsState();
    ShowState(state);
    if (backend.info)
      SetWindowTextW(Item(IDC_INFO), backend.info().c_str());
  }

  void Browse() {
    wchar_t buffer[MAX_PATH * 2] = {0};
    std::wstring current = Trim(GetText(IDC_FILE));
    wcsncpy_s(buffer, current.c_str(), _TRUNCATE);
    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter =
        L"Excel 工作簿 (*.xlsx;*.xlsm)\0*.xlsx;*.xlsm\0所有文件 (*.*)\0*.*\0";
    ofn.lpstrFile = buffer;
    ofn.nMaxFile = _countof(buffer);
    ofn.lpstrTitle = L"选择 Excel 码表";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY;
    if (GetOpenFileNameW(&ofn))
      SetWindowTextW(Item(IDC_FILE), buffer);
  }

  // true: applied (or nothing to apply)
  bool Apply() {
    SettingsState next = ReadState();
    SettingsResult result;
    if (backend.apply)
      result = backend.apply(state, next);
    else
      result.ok = true;
    if (!result.ok) {
      MessageBoxW(hwnd, result.message.c_str(), L"设置",
                  MB_OK | MB_ICONWARNING);
      return false;
    }
    state = next;
    // show the cleaned-up values (e.g. quotes around a pasted path removed)
    SetWindowTextW(Item(IDC_FILE), state.excel_file.c_str());
    SetWindowTextW(Item(IDC_SHEET), state.sheet.c_str());
    if (!result.message.empty())
      SetWindowTextW(Item(IDC_INFO), result.message.c_str());
    else if (backend.info)
      SetWindowTextW(Item(IDC_INFO), backend.info().c_str());
    return true;
  }
};

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
  SettingsWindow* w = reinterpret_cast<SettingsWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_NCCREATE: {
      auto cs = reinterpret_cast<CREATESTRUCTW*>(lp);
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, (LONG_PTR)cs->lpCreateParams);
      break;
    }
    case WM_CREATE: {
      w = reinterpret_cast<SettingsWindow*>(
          reinterpret_cast<CREATESTRUCTW*>(lp)->lpCreateParams);
      w->hwnd = hwnd;
      w->Create();
      return 0;
    }
    case WM_COMMAND:
      if (!w)
        break;
      switch (LOWORD(wp)) {
        case IDOK:
          if (w->Apply())
            DestroyWindow(hwnd);
          return 0;
        case IDCANCEL:
          DestroyWindow(hwnd);
          return 0;
        case IDC_APPLY:
          w->Apply();
          return 0;
        case IDC_BROWSE:
          w->Browse();
          return 0;
        case IDC_STATS:
          if (w->backend.show_stats)
            w->backend.show_stats();
          return 0;
      }
      break;
    case WM_CTLCOLORSTATIC:
      if (w && reinterpret_cast<HWND>(lp) == w->Item(IDC_INFO)) {
        SetTextColor(reinterpret_cast<HDC>(wp), RGB(110, 110, 110));
        SetBkMode(reinterpret_cast<HDC>(wp), TRANSPARENT);
        return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
      }
      break;
    case WM_CLOSE:
      DestroyWindow(hwnd);
      return 0;
    case WM_DESTROY:
      PostQuitMessage(0);
      return 0;
  }
  return DefWindowProcW(hwnd, msg, wp, lp);
}

void RunWindow(SettingsBackend backend) {
  SettingsWindow w;
  w.backend = std::move(backend);

  WNDCLASSEXW wc = {};
  wc.cbSize = sizeof(wc);
  wc.lpfnWndProc = WndProc;
  wc.hInstance = GetModuleHandle(NULL);
  wc.hCursor = LoadCursor(NULL, IDC_ARROW);
  wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_BTNFACE + 1);
  wc.lpszClassName = kClassName;
  RegisterClassExW(&wc);  // fails harmlessly when already registered

  HDC dc = GetDC(NULL);
  int dpi = GetDeviceCaps(dc, LOGPIXELSY);
  ReleaseDC(NULL, dc);
  if (dpi <= 0)
    dpi = 96;
  RECT rc = {0, 0, MulDiv(520, dpi, 96), MulDiv(356, dpi, 96)};
  DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU;
  DWORD ex_style = WS_EX_TOPMOST;
  AdjustWindowRectEx(&rc, style, FALSE, ex_style);
  int width = rc.right - rc.left, height = rc.bottom - rc.top;
  int x = (GetSystemMetrics(SM_CXSCREEN) - width) / 2;
  int y = (GetSystemMetrics(SM_CYSCREEN) - height) / 2;

  HWND hwnd = CreateWindowExW(ex_style, kClassName, L"小兵五笔86 设置", style, x,
                              y, width, height, NULL, NULL,
                              GetModuleHandle(NULL), &w);
  if (!hwnd) {
    std::lock_guard<std::mutex> lock(g_mutex);
    g_starting = false;
    return;
  }
  {
    std::lock_guard<std::mutex> lock(g_mutex);
    g_window = hwnd;
    g_starting = false;
  }
  ShowWindow(hwnd, SW_SHOW);
  SetForegroundWindow(hwnd);

  MSG msg;
  while (GetMessageW(&msg, NULL, 0, 0) > 0) {
    if (!IsDialogMessageW(hwnd, &msg)) {
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  }
  {
    std::lock_guard<std::mutex> lock(g_mutex);
    g_window = NULL;
  }
  if (w.font)
    DeleteObject(w.font);
}

}  // namespace

void ShowSettingsWindow(const SettingsBackend& backend) {
  {
    std::lock_guard<std::mutex> lock(g_mutex);
    if (g_window) {
      ShowWindow(g_window, SW_RESTORE);
      SetForegroundWindow(g_window);
      return;
    }
    if (g_starting)
      return;
    g_starting = true;
  }
  std::thread([backend] { RunWindow(backend); }).detach();
}

}  // namespace excel_edition
