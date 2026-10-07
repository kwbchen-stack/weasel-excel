#pragma once
//
// Excel edition: the settings window opened from the 「设置…」 menu item.
//
// The window only collects values; reading and applying them is done by the
// callbacks of SettingsBackend (see WeaselServerApp.cpp).
//
#include <functional>
#include <string>

namespace excel_edition {

struct SettingsState {
  bool pinyin = false;
  bool single_char = false;
  bool horizontal = false;
  bool completion = false;
  std::wstring excel_file;
  std::wstring sheet;
};

struct SettingsResult {
  bool ok = false;
  // ok == false: the reason, shown in a message box (nothing was applied)
  // ok == true: replaces the grey information line when not empty
  std::wstring message;
};

struct SettingsBackend {
  // current values (also called when the window opens)
  std::function<SettingsState()> load;
  // applies the changes between `before` and `after`
  std::function<SettingsResult(const SettingsState& before,
                               const SettingsState& after)>
      apply;
  // opens the typing statistics (must not block)
  std::function<void()> show_stats;
  // text of the grey line under the Excel settings
  std::function<std::wstring()> info;
};

// Opens the window on its own thread and returns at once. If the window is
// already open it is brought to the front instead.
void ShowSettingsWindow(const SettingsBackend& backend);

}  // namespace excel_edition
