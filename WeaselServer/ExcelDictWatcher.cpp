#include "stdafx.h"
#include "ExcelDictWatcher.h"
#include <shellapi.h>
#include <algorithm>
#include <fstream>
#include <sstream>

namespace fs = std::filesystem;

namespace {

// wait this long after the last change before reading the workbook
constexpr DWORD kSettleDelayMs = 1500;
// give up waiting for a workbook that stays unreadable after this many tries
constexpr int kMaxRetries = 10;
constexpr DWORD kBufferBytes = 64 * 1024;

void Trace(const std::wstring& message) {
  OutputDebugStringW((L"[ExcelDictWatcher] " + message + L"\n").c_str());
}

std::wstring Utf8ToWide(const std::string& s) {
  if (s.empty())
    return std::wstring();
  int len = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), NULL, 0);
  std::wstring w(len, L'\0');
  MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), &w[0], len);
  return w;
}

bool SameName(const std::wstring& a, const std::wstring& b) {
  return CompareStringOrdinal(a.c_str(), (int)a.size(), b.c_str(),
                              (int)b.size(), TRUE) == CSTR_EQUAL;
}

bool EndsWithNoCase(const std::wstring& s, const std::wstring& suffix) {
  return s.size() >= suffix.size() &&
         SameName(s.substr(s.size() - suffix.size()), suffix);
}

std::string Trim(const std::string& s) {
  size_t b = s.find_first_not_of(" \t\r\n");
  if (b == std::string::npos)
    return std::string();
  size_t e = s.find_last_not_of(" \t\r\n");
  return s.substr(b, e - b + 1);
}

// Unquotes a YAML scalar: 'single', "double" or plain (with # comments).
std::string YamlScalar(const std::string& raw) {
  std::string v = Trim(raw);
  if (v.empty())
    return v;
  if (v[0] == '\'') {
    std::string out;
    for (size_t i = 1; i < v.size(); ++i) {
      if (v[i] == '\'') {
        if (i + 1 < v.size() && v[i + 1] == '\'') {
          out.push_back('\'');
          ++i;
          continue;
        }
        break;
      }
      out.push_back(v[i]);
    }
    return out;
  }
  if (v[0] == '"') {
    std::string out;
    for (size_t i = 1; i < v.size(); ++i) {
      char c = v[i];
      if (c == '"')
        break;
      if (c == '\\' && i + 1 < v.size()) {
        char n = v[++i];
        out.push_back(n == 't' ? '\t' : n == 'n' ? '\n' : n);
        continue;
      }
      out.push_back(c);
    }
    return out;
  }
  size_t comment = v.find(" #");
  if (comment != std::string::npos)
    v = Trim(v.substr(0, comment));
  return v;
}

// Reads a whole file without blocking other programs (Excel keeps the
// workbook open while editing it).
bool ReadAllShared(const fs::path& file, std::string* content) {
  HANDLE h = CreateFileW(file.c_str(), GENERIC_READ,
                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                         NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
  if (h == INVALID_HANDLE_VALUE)
    return false;
  LARGE_INTEGER size;
  bool ok = GetFileSizeEx(h, &size) && size.QuadPart < 0x7FFFFFFF;
  if (ok) {
    content->resize((size_t)size.QuadPart);
    DWORD read = 0;
    ok = content->empty() ||
         (ReadFile(h, &(*content)[0], (DWORD)content->size(), &read, NULL) &&
          read == content->size());
  }
  CloseHandle(h);
  return ok;
}

uint64_t Fnv1a(const std::string& data) {
  uint64_t h = 1469598103934665603ULL;
  for (unsigned char c : data) {
    h ^= c;
    h *= 1099511628211ULL;
  }
  return h;
}

bool GetWriteTime(const fs::path& file, ULONGLONG* time) {
  WIN32_FILE_ATTRIBUTE_DATA data;
  if (!GetFileAttributesExW(file.c_str(), GetFileExInfoStandard, &data))
    return false;
  *time = ((ULONGLONG)data.ftLastWriteTime.dwHighDateTime << 32) |
          data.ftLastWriteTime.dwLowDateTime;
  return true;
}

}  // namespace

ExcelDictWatcher::ExcelDictWatcher(fs::path user_data_dir,
                                   fs::path deployer_exe)
    : user_data_dir_(std::move(user_data_dir)),
      deployer_exe_(std::move(deployer_exe)) {}

ExcelDictWatcher::~ExcelDictWatcher() {
  Stop();
}

void ExcelDictWatcher::Start() {
  if (thread_.joinable())
    return;
  stop_event_ = CreateEventW(NULL, TRUE, FALSE, NULL);
  if (!stop_event_)
    return;
  thread_ = std::thread([this] { Run(); });
}

void ExcelDictWatcher::Stop() {
  if (thread_.joinable()) {
    SetEvent(stop_event_);
    thread_.join();
  }
  if (stop_event_) {
    CloseHandle(stop_event_);
    stop_event_ = NULL;
  }
}

fs::path ExcelDictWatcher::ParseExcelSource(const fs::path& dict_file) {
  std::ifstream fin(dict_file, std::ios::in | std::ios::binary);
  if (!fin)
    return fs::path();
  std::string line, value;
  bool first = true, in_header = false, in_excel = false;
  while (std::getline(fin, line)) {
    if (first && line.size() >= 3 && line.compare(0, 3, "\xEF\xBB\xBF") == 0)
      line = line.substr(3);  // UTF-8 BOM
    first = false;
    std::string t = Trim(line);
    if (t == "---") {
      in_header = true;
      continue;
    }
    if (t == "...")
      break;
    if (t.empty() || t[0] == '#')
      continue;
    bool indented = line[0] == ' ' || line[0] == '\t';
    if (!indented) {
      in_header = true;  // header may omit the leading ---
      in_excel = t.compare(0, 6, "excel:") == 0;
      continue;
    }
    if (in_header && in_excel && t.compare(0, 5, "file:") == 0) {
      value = YamlScalar(t.substr(5));
      break;
    }
  }
  if (value.empty())
    return fs::path();
  fs::path excel(Utf8ToWide(value));
  if (excel.is_relative())
    excel = dict_file.parent_path() / excel;
  return excel.lexically_normal();
}

void ExcelDictWatcher::ScanSources() {
  std::vector<Source> old = std::move(sources_);
  sources_.clear();
  std::error_code ec;
  for (fs::directory_iterator it(user_data_dir_, ec), end; !ec && it != end;
       it.increment(ec)) {
    const fs::path& p = it->path();
    std::wstring name = p.filename().wstring();
    if (!EndsWithNoCase(name, L".dict.yaml"))
      continue;
    fs::path excel = ParseExcelSource(p);
    if (excel.empty())
      continue;
    Source s;
    s.excel_file = excel;
    {
      // settings in the .dict.yaml (e.g. the sheet name) are part of what the
      // table is built from: when they change, deploy again even though the
      // workbook itself did not
      std::string dict_content;
      if (ReadAllShared(p, &dict_content))
        s.dict_hash = Fnv1a(dict_content);
    }
    std::wstring dict_name = name.substr(0, name.size() - 10);  // .dict.yaml
    s.table_file = user_data_dir_ / L"build" / (dict_name + L".table.bin");
    for (const auto& o : old) {
      if (SameName(o.excel_file.wstring(), excel.wstring()) &&
          o.dict_hash == s.dict_hash) {
        s.deployed_hash = o.deployed_hash;
        s.has_hash = o.has_hash;
      }
    }
    Trace(L"watching " + excel.wstring() + L" for " + name);
    sources_.push_back(std::move(s));
  }
}

bool ExcelDictWatcher::Arm(Folder& folder) {
  ResetEvent(folder.event);
  folder.overlapped = {};
  folder.overlapped.hEvent = folder.event;
  return ReadDirectoryChangesW(
             folder.handle, folder.buffer.data(),
             (DWORD)(folder.buffer.size() * sizeof(DWORD)), FALSE,
             FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_LAST_WRITE |
                 FILE_NOTIFY_CHANGE_SIZE,
             NULL, &folder.overlapped, NULL) != FALSE;
}

void ExcelDictWatcher::OpenFolders() {
  auto folder_for = [this](const fs::path& dir) -> Folder* {
    for (auto& f : folders_) {
      if (SameName(f->dir.wstring(), dir.wstring()))
        return f.get();
    }
    auto f = std::make_unique<Folder>();
    f->dir = dir;
    folders_.push_back(std::move(f));
    return folders_.back().get();
  };
  folder_for(user_data_dir_)->watch_dict_yaml = true;
  for (const auto& s : sources_) {
    folder_for(s.excel_file.parent_path())
        ->excel_names.push_back(s.excel_file.filename().wstring());
  }
  for (auto& f : folders_) {
    f->handle = CreateFileW(
        f->dir.c_str(), FILE_LIST_DIRECTORY,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, NULL,
        OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED, NULL);
    f->event = CreateEventW(NULL, TRUE, FALSE, NULL);
    f->buffer.resize(kBufferBytes / sizeof(DWORD));
    if (f->handle == INVALID_HANDLE_VALUE || !f->event || !Arm(*f)) {
      Trace(L"cannot watch folder " + f->dir.wstring());
      if (f->handle != INVALID_HANDLE_VALUE)
        CloseHandle(f->handle);
      f->handle = INVALID_HANDLE_VALUE;
    }
  }
}

void ExcelDictWatcher::CloseFolders() {
  for (auto& f : folders_) {
    if (f->handle != INVALID_HANDLE_VALUE) {
      CancelIoEx(f->handle, &f->overlapped);
      DWORD ignored;
      GetOverlappedResult(f->handle, &f->overlapped, &ignored, TRUE);
      CloseHandle(f->handle);
    }
    if (f->event)
      CloseHandle(f->event);
  }
  folders_.clear();
}

bool ExcelDictWatcher::HandleNotifications(Folder& folder, DWORD bytes) {
  bool relevant = false;
  if (bytes == 0) {
    // buffer overflow: changes were lost, check everything in this folder
    if (!folder.excel_names.empty())
      excel_pending_ = relevant = true;
    if (folder.watch_dict_yaml)
      rescan_pending_ = relevant = true;
    return relevant;
  }
  const BYTE* p = reinterpret_cast<const BYTE*>(folder.buffer.data());
  while (true) {
    auto info = reinterpret_cast<const FILE_NOTIFY_INFORMATION*>(p);
    std::wstring name(info->FileName, info->FileNameLength / sizeof(WCHAR));
    if (folder.watch_dict_yaml && EndsWithNoCase(name, L".dict.yaml"))
      rescan_pending_ = relevant = true;
    for (const auto& target : folder.excel_names) {
      if (SameName(name, target))
        excel_pending_ = relevant = true;
    }
    if (info->NextEntryOffset == 0)
      break;
    p += info->NextEntryOffset;
  }
  return relevant;
}

bool ExcelDictWatcher::CheckSources(bool at_startup) {
  bool need_deploy = false;
  bool retry = false;
  std::vector<uint64_t> hashes(sources_.size(), 0);
  std::vector<bool> readable(sources_.size(), false);
  for (size_t i = 0; i < sources_.size(); ++i) {
    Source& s = sources_[i];
    std::string content;
    if (!ReadAllShared(s.excel_file, &content) || content.size() < 4 ||
        content.compare(0, 2, "PK") != 0) {
      // missing, locked or half written (Excel saves in several steps)
      retry = true;
      continue;
    }
    readable[i] = true;
    hashes[i] = Fnv1a(content);
    if (at_startup) {
      // changed while Weasel was not running?
      ULONGLONG excel_time = 0, table_time = 0;
      bool newer = !GetWriteTime(s.table_file, &table_time) ||
                   (GetWriteTime(s.excel_file, &excel_time) &&
                    excel_time > table_time);
      if (newer)
        need_deploy = true;
    } else if (!s.has_hash || hashes[i] != s.deployed_hash) {
      need_deploy = true;
    }
  }
  // a workbook is still being written: check everything again later,
  // without recording anything yet
  if (retry && !at_startup && retries_ < kMaxRetries)
    return true;
  bool deployed = !need_deploy || Deploy();
  for (size_t i = 0; i < sources_.size(); ++i) {
    if (!readable[i])
      continue;
    // after a failed start of the deployer, the next change must retry it
    sources_[i].deployed_hash = hashes[i];
    sources_[i].has_hash = deployed;
  }
  return false;
}

bool ExcelDictWatcher::Deploy() {
  // wait for a deployment that is already running (e.g. started from the
  // tray menu); WeaselDeployer holds this mutex while it works.
  HANDLE mutex = OpenMutexW(SYNCHRONIZE | MUTEX_MODIFY_STATE, FALSE,
                            L"WeaselDeployerMutex");
  if (mutex) {
    HANDLE waits[] = {stop_event_, mutex};
    DWORD r = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
    if (r == WAIT_OBJECT_0 + 1 || r == WAIT_ABANDONED_0 + 1)
      ReleaseMutex(mutex);
    CloseHandle(mutex);
    if (r == WAIT_OBJECT_0)
      return false;  // stopping
  }
  Trace(L"workbook changed, deploying");
  SHELLEXECUTEINFOW info = {};
  info.cbSize = sizeof(info);
  info.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
  info.lpFile = deployer_exe_.c_str();
  info.lpParameters = L"/deploy";
  info.nShow = SW_SHOWNORMAL;
  if (!ShellExecuteExW(&info) || !info.hProcess) {
    Trace(L"failed to start " + deployer_exe_.wstring());
    return false;
  }
  // wait until it finishes so that deployments never overlap
  HANDLE waits[] = {stop_event_, info.hProcess};
  WaitForMultipleObjects(2, waits, FALSE, INFINITE);
  CloseHandle(info.hProcess);
  return true;
}

void ExcelDictWatcher::Run() {
  ScanSources();
  OpenFolders();
  // catch up with changes made while Weasel was not running
  CheckSources(true);

  while (true) {
    std::vector<HANDLE> handles{stop_event_};
    std::vector<Folder*> owners{nullptr};
    for (auto& f : folders_) {
      if (f->handle != INVALID_HANDLE_VALUE) {
        handles.push_back(f->event);
        owners.push_back(f.get());
      }
    }
    DWORD timeout = INFINITE;
    if (excel_pending_ || rescan_pending_) {
      ULONGLONG now = GetTickCount64();
      timeout = deadline_ > now ? (DWORD)(deadline_ - now) : 0;
    }
    DWORD r = WaitForMultipleObjects((DWORD)handles.size(), handles.data(),
                                     FALSE, timeout);
    if (r == WAIT_OBJECT_0 || r == WAIT_FAILED)
      break;  // stopping
    if (r == WAIT_TIMEOUT) {
      if (rescan_pending_) {
        rescan_pending_ = false;
        CloseFolders();
        ScanSources();
        OpenFolders();
        excel_pending_ = true;  // a new or moved workbook may need deploying
      }
      if (excel_pending_) {
        excel_pending_ = false;
        if (CheckSources(false)) {
          ++retries_;
          excel_pending_ = true;
          deadline_ = GetTickCount64() + kSettleDelayMs;
        } else {
          retries_ = 0;
        }
      }
      continue;
    }
    DWORD index = r - WAIT_OBJECT_0;
    if (index >= owners.size())
      continue;
    Folder* folder = owners[index];
    DWORD bytes = 0;
    bool relevant = false;
    if (GetOverlappedResult(folder->handle, &folder->overlapped, &bytes,
                            FALSE)) {
      relevant = HandleNotifications(*folder, bytes);
    }
    if (!Arm(*folder)) {
      // the folder went away (e.g. renamed); pick it up again on rescan
      Trace(L"lost watch on " + folder->dir.wstring());
      CloseHandle(folder->handle);
      folder->handle = INVALID_HANDLE_VALUE;
    }
    if (relevant) {
      retries_ = 0;
      // restart the settle delay on every change to a watched file
      deadline_ = GetTickCount64() + kSettleDelayMs;
    }
  }
  CloseFolders();
}
