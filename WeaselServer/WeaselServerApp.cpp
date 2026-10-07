#include "stdafx.h"
#include "WeaselServerApp.h"
#include <filesystem>
#include <fstream>
#include <iterator>

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
  // Excel edition: these commands also arrive from the language bar menu of
  // WeaselTSF, forwarded as tray commands over the IPC pipe
  using namespace excel_edition;
  m_server.AddMenuHandler(ID_PINYIN_MODE, [this] {
    m_server.WithApiLock(
        [this] { m_handler->SetPinyinMode(!m_handler->IsPinyinMode()); });
    return true;
  });
  m_server.AddMenuHandler(ID_SINGLE_CHAR, [this] {
    m_server.WithApiLock(
        [this] { m_handler->SetSingleChar(!m_handler->IsSingleChar()); });
    return true;
  });
  m_server.AddMenuHandler(ID_HORIZONTAL, [this] {
    m_server.WithApiLock(
        [this] { m_handler->SetHorizontal(!m_handler->IsHorizontal()); });
    return true;
  });
  m_server.AddMenuHandler(ID_COMPLETION, [this] {
    m_server.WithApiLock(
        [this] { m_handler->SetCompletion(!m_handler->IsCompletion()); });
    return true;
  });
  m_server.AddMenuHandler(ID_SETTINGS, [this] {
    // the window runs on its own thread (see SettingsDialog.cpp)
    OpenSettings();
    return true;
  });
  m_server.AddMenuHandler(ID_TYPING_STATS, [this] {
    // show the dialog on its own thread: a request from the language bar is
    // handled while holding the API lock, and typing must not wait for the
    // dialog to be closed
    std::thread([this] { ShowTypingStats(); }).detach();
    return true;
  });
}

void WeaselServerApp::CustomizeTrayMenu(HMENU menu) {
  bool pinyin = false;
  m_server.WithApiLock([&] { pinyin = m_handler->IsPinyinMode(); });
  excel_edition::InsertMenuItems(menu, pinyin);
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

namespace {

std::wstring Utf8ToW(const std::string& s) {
  if (s.empty())
    return std::wstring();
  int len = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), NULL, 0);
  std::wstring w(len, L'\0');
  MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), &w[0], len);
  return w;
}

std::string WToUtf8(const std::wstring& w) {
  if (w.empty())
    return std::string();
  int len = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), NULL, 0,
                                NULL, NULL);
  std::string s(len, '\0');
  WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), &s[0], len, NULL,
                      NULL);
  return s;
}

bool ReadTextFile(const fs::path& file, std::string* content) {
  std::ifstream fin(file, std::ios::in | std::ios::binary);
  if (!fin)
    return false;
  content->assign(std::istreambuf_iterator<char>(fin),
                  std::istreambuf_iterator<char>());
  return true;
}

}  // namespace

// ---- settings window ----

void WeaselServerApp::OpenSettings() {
  using namespace excel_edition;
  const fs::path dict_file = WeaselUserDataPath() / kWubiDictFile;
  const fs::path table_file = WeaselUserDataPath() / L"build" / kWubiTableFile;

  SettingsBackend backend;

  backend.load = [this, dict_file]() {
    SettingsState s;
    m_server.WithApiLock([&] {
      s.pinyin = m_handler->IsPinyinMode();
      s.single_char = m_handler->IsSingleChar();
      s.horizontal = m_handler->IsHorizontal();
      s.completion = m_handler->IsCompletion();
    });
    std::string text;
    ExcelSettings excel;
    if (ReadTextFile(dict_file, &text) && ParseExcelSettings(text, &excel)) {
      s.excel_file = Utf8ToW(excel.file);
      s.sheet = Utf8ToW(excel.sheet);
    }
    return s;
  };

  backend.info = [table_file]() -> std::wstring {
    uint32_t entries = 0;
    if (!ReadTableEntryCount(table_file, &entries))
      return L"还没有生成词库。";
    std::wstring text = L"上次生成的词库：" + FormatCount(entries) + L" 条";
    WIN32_FILE_ATTRIBUTE_DATA data;
    if (GetFileAttributesExW(table_file.c_str(), GetFileExInfoStandard,
                             &data)) {
      FILETIME local;
      SYSTEMTIME t;
      if (FileTimeToLocalFileTime(&data.ftLastWriteTime, &local) &&
          FileTimeToSystemTime(&local, &t)) {
        wchar_t buf[64];
        swprintf(buf, _countof(buf), L"（%04d-%02d-%02d %02d:%02d）", t.wYear,
                 t.wMonth, t.wDay, t.wHour, t.wMinute);
        text += buf;
      }
    }
    return text;
  };

  backend.show_stats = [this] {
    // the statistics dialog blocks: give it a thread of its own
    std::thread([this] { ShowTypingStats(); }).detach();
  };

  backend.apply = [this, dict_file](const SettingsState& before,
                                    const SettingsState& after) {
    SettingsResult result;
    bool excel_changed =
        before.excel_file != after.excel_file || before.sheet != after.sheet;

    // 1. check everything that can be wrong before changing anything
    std::string new_dict_text;
    if (excel_changed) {
      if (after.excel_file.empty()) {
        result.message = L"请填写 Excel 文件的路径。";
        return result;
      }
      if (after.sheet.empty()) {
        result.message = L"请填写工作表名称。";
        return result;
      }
      fs::path file(after.excel_file);
      switch (ClassifyWorkbookName(WToUtf8(file.filename().wstring()))) {
        case WorkbookKind::kOldXls:
          result.message =
              L"不支持 .xls 格式（老版本 Excel 格式，一张表最多 65,536 行）。\n"
              L"请先在 Excel 里「另存为」.xlsx 格式，再选择新文件。";
          return result;
        case WorkbookKind::kOther:
          result.message = L"只支持 .xlsx 格式的 Excel 文件。";
          return result;
        case WorkbookKind::kSupported:
          break;
      }
      std::error_code ec;
      if (!fs::is_regular_file(file, ec)) {
        result.message = L"找不到这个文件：\n" + after.excel_file;
        return result;
      }
      std::string dict_text;
      if (!ReadTextFile(dict_file, &dict_text)) {
        result.message = L"读不到码表设置文件：\n" + dict_file.wstring();
        return result;
      }
      ExcelSettings settings{WToUtf8(after.excel_file), WToUtf8(after.sheet)};
      if (!ReplaceExcelSettings(dict_text, settings, &new_dict_text)) {
        result.message =
            L"码表设置文件里没有找到 excel: 一段，不能修改：\n" +
            dict_file.wstring();
        return result;
      }
    }

    // 2. the Excel settings: back up the settings file, write the new one next
    // to it and swap, so that the watcher never sees a half written file
    if (excel_changed) {
      std::error_code ec;
      fs::path backup = dict_file;
      backup += L".bak";
      fs::copy_file(dict_file, backup, fs::copy_options::overwrite_existing,
                    ec);
      fs::path temp = dict_file;
      temp += L".tmp";
      {
        std::ofstream fout(temp, std::ios::out | std::ios::binary);
        fout.write(new_dict_text.data(), (std::streamsize)new_dict_text.size());
        if (!fout) {
          result.message = L"写入码表设置文件失败：\n" + temp.wstring();
          return result;
        }
      }
      fs::rename(temp, dict_file, ec);
      if (ec) {
        result.message = L"保存码表设置文件失败：\n" + dict_file.wstring();
        return result;
      }
    }

    // 3. the switches
    m_server.WithApiLock([&] {
      if (before.pinyin != after.pinyin)
        m_handler->SetPinyinMode(after.pinyin);
      if (before.single_char != after.single_char)
        m_handler->SetSingleChar(after.single_char);
      if (before.horizontal != after.horizontal)
        m_handler->SetHorizontal(after.horizontal);
      if (before.completion != after.completion)
        m_handler->SetCompletion(after.completion);
    });

    result.ok = true;
    if (excel_changed)
      result.message =
          L"已保存，正在重新生成词库（几秒钟）。生成完后重新打开本窗口，可看到新的词条数。";
    return result;
  };

  ShowSettingsWindow(backend);
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
