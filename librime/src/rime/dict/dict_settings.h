//
// Copyright RIME Developers
// Distributed under the BSD License
//
// 2012-11-11 GONG Chen <chen.sst@gmail.com>
//
#ifndef RIME_DICT_SETTINGS_H_
#define RIME_DICT_SETTINGS_H_

#include <istream>
#include <rime/common.h>
#include <rime/config.h>

namespace rime {

class DictSettings : public Config {
 public:
  DictSettings();
  bool LoadDictHeader(std::istream& stream);
  bool empty();
  string dict_name();
  string dict_version();
  string sort_order();
  bool use_preset_vocabulary();
  string vocabulary();
  bool use_rule_based_encoder();
  int max_phrase_length();
  double min_phrase_weight();
  an<ConfigList> GetTables();
  int GetColumnIndex(const string& column_label);
  // dictionary entries loaded from an Excel workbook (.xlsx)
  bool has_excel_source();
  string excel_file();
  string excel_sheet();
  string excel_layout();
  int excel_header_rows();
  // resolves excel/file; relative paths are relative to the .dict.yaml file
  path ExcelFilePath(const path& dict_file);
};

}  // namespace rime

#endif  // RIME_DICT_SETTINGS_H_
