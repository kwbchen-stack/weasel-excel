#pragma once
#include <WeaselUI.h>
#include <WeaselIPC.h>
#include "SystemTraySDK.h"
#include <functional>

#define WM_WEASEL_TRAY_NOTIFY (WEASEL_IPC_LAST_COMMAND + 100)

class WeaselTrayIcon : public CSystemTray {
 public:
  enum WeaselTrayMode {
    INITIAL,
    ZHUNG,
    ASCII,
    DISABLED,
  };

  WeaselTrayIcon(weasel::UI& ui);

  BOOL Create(HWND hTargetWnd);
  void Refresh();
  // called right before the tray menu pops up, to add or check items
  void SetMenuCustomizer(std::function<void(HMENU)> customizer) {
    m_menu_customizer = std::move(customizer);
  }

 protected:
  virtual void CustomizeMenu(HMENU hMenu);

  weasel::UIStyle& m_style;
  weasel::Status& m_status;
  WeaselTrayMode m_mode;
  std::wstring m_schema_zhung_icon;
  std::wstring m_schema_ascii_icon;
  bool m_disabled;
  std::function<void(HMENU)> m_menu_customizer;
};
