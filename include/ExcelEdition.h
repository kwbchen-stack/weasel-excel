#pragma once
//
// Excel edition (小兵五笔86) additions shared by WeaselServer, RimeWithWeasel
// and WeaselTSF: menu command ids, registry flags and the menu items.
//
// The server applies the toggles and stores their state under
// HKCU\Software\Rime\Weasel\ExcelEdition; the language bar menu in WeaselTSF
// reads it from there to show check marks and forwards the commands to the
// server as tray commands.
//
#include <windows.h>

namespace excel_edition {

// menu command ids (must not clash with ID_WEASELTRAY_* in resource.h)
constexpr UINT ID_PINYIN_MODE = 41101;
constexpr UINT ID_SINGLE_CHAR = 41102;
constexpr UINT ID_TYPING_STATS = 41103;
constexpr UINT ID_HORIZONTAL = 41105;

// schemas and options
constexpr char kWubiSchema[] = "xiaobing_wubi86";
constexpr char kPinyinSchema[] = "xiaobing_pinyin";
constexpr char kSingleCharOption[] = "single_char";

// registry
constexpr wchar_t kRegKey[] = L"Software\\Rime\\Weasel\\ExcelEdition";
constexpr wchar_t kPinyinModeValue[] = L"PinyinMode";
constexpr wchar_t kSingleCharValue[] = L"SingleChar";
// effective layout, written by the server for the language bar menu
constexpr wchar_t kHorizontalValue[] = L"Horizontal";
// layout chosen in the menu; absent = follow weasel.yaml
constexpr wchar_t kLayoutOverrideValue[] = L"LayoutOverride";

inline bool ReadFlag(const wchar_t* name, DWORD* value) {
  DWORD size = sizeof(*value);
  return RegGetValueW(HKEY_CURRENT_USER, kRegKey, name, RRF_RT_REG_DWORD, NULL,
                      value, &size) == ERROR_SUCCESS;
}

inline bool LoadFlag(const wchar_t* name) {
  DWORD value = 0;
  return ReadFlag(name, &value) && value != 0;
}

inline void SaveFlag(const wchar_t* name, bool on) {
  DWORD value = on ? 1 : 0;
  RegSetKeyValueW(HKEY_CURRENT_USER, kRegKey, name, REG_DWORD, &value,
                  sizeof(value));
}

// Inserts the edition's items (and a separator) at the top of a popup menu.
inline void InsertMenuItems(HMENU menu,
                            bool pinyin_mode,
                            bool single_char,
                            bool horizontal) {
  UINT pos = 0;
  auto check = [](bool on) { return on ? MF_CHECKED : MF_UNCHECKED; };
  // 拼音模式 / 单字模式 / 候选横排 / 打字统计…
  InsertMenuW(menu, pos++, MF_BYPOSITION | MF_STRING | check(pinyin_mode),
              ID_PINYIN_MODE, L"拼音模式 (&P)");
  InsertMenuW(menu, pos++, MF_BYPOSITION | MF_STRING | check(single_char),
              ID_SINGLE_CHAR, L"单字模式 (&W)");
  InsertMenuW(menu, pos++, MF_BYPOSITION | MF_STRING | check(horizontal),
              ID_HORIZONTAL, L"候选横排 (&H)");
  InsertMenuW(menu, pos++, MF_BYPOSITION | MF_STRING, ID_TYPING_STATS,
              L"打字统计… (&T)");
  InsertMenuW(menu, pos++, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
}

// state as last stored by the server (used by the language bar menu)
inline void InsertMenuItemsFromRegistry(HMENU menu) {
  InsertMenuItems(menu, LoadFlag(kPinyinModeValue), LoadFlag(kSingleCharValue),
                  LoadFlag(kHorizontalValue));
}

}  // namespace excel_edition
