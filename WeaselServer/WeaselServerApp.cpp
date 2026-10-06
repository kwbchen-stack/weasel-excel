#include "stdafx.h"
#include "WeaselServerApp.h"
#include <filesystem>

WeaselServerApp::WeaselServerApp()
    : m_handler(std::make_unique<RimeWithWeaselHandler>(&m_ui)),
      tray_icon(m_ui) {
  // m_handler.reset(new RimeWithWeaselHandler(&m_ui));
  m_server.SetRequestHandler(m_handler.get());
  SetupMenuHandlers();
}

WeaselServerApp::~WeaselServerApp() {}

int WeaselServerApp::Run() {
  if (!m_server.Start())
    return -1;

  // win_sparkle_set_appcast_url("http://localhost:8000/weasel/update/appcast.xml");
  win_sparkle_set_registry_path("Software\\Rime\\Weasel\\Updates");
  if (GetThreadUILanguage() ==
      MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL))
    win_sparkle_set_lang("zh-TW");
  else if (GetThreadUILanguage() ==
           MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED))
    win_sparkle_set_lang("zh-CN");
  else
    win_sparkle_set_lang("en");
  win_sparkle_init();
  m_ui.Create(m_server.GetHWnd());

  m_handler->Initialize();
  m_handler->OnUpdateUI([this]() { tray_icon.Refresh(); });

  tray_icon.Create(m_server.GetHWnd());
  tray_icon.SetMenuCustomizer([this](HMENU menu) { CustomizeTrayMenu(menu); });
  tray_icon.Refresh();

  m_excel_watcher = std::make_unique<ExcelDictWatcher>(
      WeaselUserDataPath(), install_dir() / L"WeaselDeployer.exe");
  m_excel_watcher->Start();

  int ret = m_server.Run();

  m_excel_watcher->Stop();
  m_handler->Finalize();
  m_ui.Destroy();
  tray_icon.RemoveIcon();
  win_sparkle_cleanup();

  return ret;
}

void WeaselServerApp::SetupMenuHandlers() {
  std::filesystem::path dir = install_dir();
  m_server.AddMenuHandler(ID_WEASELTRAY_QUIT,
                          [this] { return m_server.Stop() == 0; });
  m_server.AddMenuHandler(ID_WEASELTRAY_DEPLOY,
                          std::bind(execute, dir / L"WeaselDeployer.exe",
                                    std::wstring(L"/deploy")));
  m_server.AddMenuHandler(
      ID_WEASELTRAY_SETTINGS,
      std::bind(execute, dir / L"WeaselDeployer.exe", std::wstring()));
  m_server.AddMenuHandler(
      ID_WEASELTRAY_DICT_MANAGEMENT,
      std::bind(execute, dir / L"WeaselDeployer.exe", std::wstring(L"/dict")));
  m_server.AddMenuHandler(
      ID_WEASELTRAY_SYNC,
      std::bind(execute, dir / L"WeaselDeployer.exe", std::wstring(L"/sync")));
  m_server.AddMenuHandler(ID_WEASELTRAY_WIKI,
                          std::bind(open, L"https://rime.im/docs/"));
  m_server.AddMenuHandler(ID_WEASELTRAY_HOMEPAGE,
                          std::bind(open, L"https://rime.im/"));
  m_server.AddMenuHandler(ID_WEASELTRAY_FORUM,
                          std::bind(open, L"https://rime.im/discuss/"));
  m_server.AddMenuHandler(ID_WEASELTRAY_CHECKUPDATE, check_update);
  m_server.AddMenuHandler(ID_WEASELTRAY_INSTALLDIR, std::bind(explore, dir));
  m_server.AddMenuHandler(ID_WEASELTRAY_USERCONFIG,
                          std::bind(explore, WeaselUserDataPath()));
  m_server.AddMenuHandler(ID_WEASELTRAY_LOGDIR,
                          std::bind(explore, WeaselLogPath()));
  // Excel edition
  m_server.AddMenuHandler(ID_EXCEL_MIX_MODE, [this] {
    m_server.WithApiLock(
        [this] { m_handler->SetMixMode(!m_handler->IsMixMode()); });
    return true;
  });
  m_server.AddMenuHandler(ID_EXCEL_SINGLE_CHAR, [this] {
    m_server.WithApiLock(
        [this] { m_handler->SetSingleChar(!m_handler->IsSingleChar()); });
    return true;
  });
  m_server.AddMenuHandler(ID_EXCEL_TYPING_STATS,
                          [this] { return ShowTypingStats(); });
}

void WeaselServerApp::CustomizeTrayMenu(HMENU menu) {
  bool mix = false, single_char = false;
  m_server.WithApiLock([&] {
    mix = m_handler->IsMixMode();
    single_char = m_handler->IsSingleChar();
  });
  UINT pos = 0;
  InsertMenuW(menu, pos++,
              MF_BYPOSITION | MF_STRING | (mix ? MF_CHECKED : MF_UNCHECKED),
              ID_EXCEL_MIX_MODE, L"五笔拼音混输 (&M)");
  InsertMenuW(
      menu, pos++,
      MF_BYPOSITION | MF_STRING | (single_char ? MF_CHECKED : MF_UNCHECKED),
      ID_EXCEL_SINGLE_CHAR, L"单字模式 (&W)");
  InsertMenuW(menu, pos++, MF_BYPOSITION | MF_STRING, ID_EXCEL_TYPING_STATS,
              L"打字统计… (&T)");
  InsertMenuW(menu, pos++, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
}

static std::wstring FormatCount(long long n) {
  std::wstring digits = std::to_wstring(n < 0 ? -n : n);
  std::wstring out;
  int count = 0;
  for (auto it = digits.rbegin(); it != digits.rend(); ++it) {
    if (count && count % 3 == 0)
      out.insert(out.begin(), L',');
    out.insert(out.begin(), *it);
    ++count;
  }
  return n < 0 ? L"-" + out : out;
}

bool WeaselServerApp::ShowTypingStats() {
  RimeWithWeaselHandler::TypingStats stats;
  fs::path file;
  m_server.WithApiLock([&] {
    stats = m_handler->GetTypingStats();
    file = m_handler->TypingStatsFile();
  });
  long long average = stats.days ? stats.total_han / stats.days : 0;
  std::wstring msg = L"今日：" + FormatCount(stats.today_han) + L" 字\n" +
                     L"本月：" + FormatCount(stats.month_han) + L" 字\n" +
                     L"累计：" + FormatCount(stats.total_han) + L" 字（共 " +
                     std::to_wstring(stats.days) + L" 天，日均 " +
                     FormatCount(average) + L" 字）\n\n" +
                     L"今日其他字符（标点、英文等）：" +
                     FormatCount(stats.today_other) + L" 个\n\n" +
                     L"明细表：" + file.wstring();
  std::error_code ec;
  if (!fs::exists(file, ec)) {
    MessageBoxW(NULL, (msg + L"\n（还没有记录）").c_str(), L"打字统计",
                MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND);
    return true;
  }
  msg += L"\n\n要用 Excel 打开明细表吗？\n（打开的是一份副本，不影响继续统计）";
  if (MessageBoxW(NULL, msg.c_str(), L"打字统计",
                  MB_YESNO | MB_ICONINFORMATION | MB_SETFOREGROUND) == IDYES) {
    WCHAR temp_dir[MAX_PATH] = {0};
    GetTempPathW(MAX_PATH, temp_dir);
    fs::path copy = fs::path(temp_dir) / L"打字统计明细（副本）.csv";
    if (CopyFileW(file.c_str(), copy.c_str(), FALSE))
      open(copy);
  }
  return true;
}
