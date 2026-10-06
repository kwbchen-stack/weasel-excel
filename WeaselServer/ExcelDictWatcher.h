#pragma once
//
// ExcelDictWatcher: redeploys Rime when an Excel workbook used as a
// dictionary source (see `excel/file` in *.dict.yaml) is saved.
//
// Fully event-driven: the worker thread sleeps in WaitForMultipleObjects
// until Windows reports a change in a watched folder (ReadDirectoryChangesW).
// The only timer is a one-shot "settle" delay armed by such an event, so that
// the several file operations of a single Excel save trigger one deployment.
//
#include <windows.h>
#include <cstdint>
#include <filesystem>
#include <map>
#include <memory>
#include <string>
#include <thread>
#include <vector>

class ExcelDictWatcher {
 public:
  ExcelDictWatcher(std::filesystem::path user_data_dir,
                   std::filesystem::path deployer_exe);
  ~ExcelDictWatcher();

  void Start();
  void Stop();

  // Returns the workbook referenced by `excel/file` in the header of a
  // .dict.yaml file, or an empty path if there is none.
  static std::filesystem::path ParseExcelSource(
      const std::filesystem::path& dict_file);

 private:
  struct Source {
    std::filesystem::path excel_file;  // absolute path of the workbook
    std::filesystem::path table_file;  // build\<dict>.table.bin
    uint64_t deployed_hash = 0;        // content hash at last deployment
    bool has_hash = false;
  };
  struct Folder {
    std::filesystem::path dir;
    HANDLE handle = INVALID_HANDLE_VALUE;
    HANDLE event = NULL;
    OVERLAPPED overlapped = {};
    std::vector<DWORD> buffer;  // DWORD-aligned as required by the API
    bool watch_dict_yaml = false;
    std::vector<std::wstring> excel_names;  // file names inside this folder
  };

  void Run();
  void ScanSources();
  void OpenFolders();
  void CloseFolders();
  bool Arm(Folder& folder);
  // returns true if a watched file was touched
  bool HandleNotifications(Folder& folder, DWORD bytes);
  // returns true if a check must be retried later (file still being written)
  bool CheckSources(bool at_startup);
  bool Deploy();

  std::filesystem::path user_data_dir_;
  std::filesystem::path deployer_exe_;
  std::vector<Source> sources_;
  std::vector<std::unique_ptr<Folder>> folders_;
  HANDLE stop_event_ = NULL;
  std::thread thread_;
  bool excel_pending_ = false;
  bool rescan_pending_ = false;
  int retries_ = 0;
  ULONGLONG deadline_ = 0;
};
