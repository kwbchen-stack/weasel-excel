//
// Copyright RIME Developers
// Distributed under the BSD License
//
// Minimal reader for Excel .xlsx workbooks (Office Open XML), used to
// load dictionary entries directly from a spreadsheet.
//
#ifndef RIME_XLSX_READER_H_
#define RIME_XLSX_READER_H_

#include <rime/common.h>

namespace rime {

struct XlsxRow {
  int row_number = 0;     // 1-based row number as shown in Excel
  vector<string> cells;   // cells[0] is column A; missing cells are ""
};

class XlsxReader {
 public:
  // Reads the whole workbook file into memory (shared read, binary mode).
  bool Open(const path& file_path);
  // Reads all rows of the named sheet; an empty name selects the first sheet.
  bool ReadSheet(const string& sheet_name, vector<XlsxRow>* rows);
  const string& error() const { return error_; }

  // Reads a file in binary mode. Returns false if it cannot be opened.
  static bool ReadFileBytes(const path& file_path, string* content);

 private:
  bool ExtractEntry(const string& entry_name, string* content, bool required);
  bool LoadSharedStrings();

  string data_;
  vector<string> shared_strings_;
  string error_;
};

}  // namespace rime

#endif  // RIME_XLSX_READER_H_
