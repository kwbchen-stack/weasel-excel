#pragma once
//
// Excel edition: reading and rewriting the `excel:` block of a *.dict.yaml,
// and reading the entry count of a compiled table.bin.
//
// Pure C++17 (no Windows headers) so that it can be tested anywhere.
// Text is UTF-8.
//
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <string>
#include <vector>

namespace excel_edition {

struct ExcelSettings {
  std::string file;   // path of the workbook
  std::string sheet;  // sheet name
};

namespace detail {

inline std::string Trim(const std::string& s) {
  size_t b = 0, e = s.size();
  while (b < e && (s[b] == ' ' || s[b] == '\t' || s[b] == '\r' || s[b] == '\n'))
    ++b;
  while (e > b &&
         (s[e - 1] == ' ' || s[e - 1] == '\t' || s[e - 1] == '\r' ||
          s[e - 1] == '\n'))
    --e;
  return s.substr(b, e - b);
}

// value of a one line YAML scalar (plain, 'single' or "double" quoted);
// a trailing  # comment  after a plain scalar is dropped
inline std::string Scalar(const std::string& raw) {
  std::string s = Trim(raw);
  if (s.size() >= 2 && s.front() == '\'') {
    std::string out;
    for (size_t i = 1; i < s.size(); ++i) {
      if (s[i] == '\'') {
        if (i + 1 < s.size() && s[i + 1] == '\'') {
          out += '\'';
          ++i;
        } else {
          break;
        }
      } else {
        out += s[i];
      }
    }
    return out;
  }
  if (s.size() >= 2 && s.front() == '"') {
    std::string out;
    for (size_t i = 1; i < s.size(); ++i) {
      if (s[i] == '\\' && i + 1 < s.size()) {
        out += s[++i];
      } else if (s[i] == '"') {
        break;
      } else {
        out += s[i];
      }
    }
    return out;
  }
  size_t hash = s.find(" #");
  if (hash != std::string::npos)
    s = Trim(s.substr(0, hash));
  return s;
}

inline std::string Quote(const std::string& value) {
  std::string out = "'";
  for (char c : value) {
    if (c == '\'')
      out += "''";
    else
      out += c;
  }
  out += '\'';
  return out;
}

struct Line {
  size_t begin = 0, end = 0;  // [begin, end) without the line break
  size_t next = 0;            // start of the next line
};

inline std::vector<Line> SplitLines(const std::string& text) {
  std::vector<Line> lines;
  size_t pos = 0;
  while (pos < text.size()) {
    Line l;
    l.begin = pos;
    size_t nl = text.find('\n', pos);
    if (nl == std::string::npos) {
      l.end = text.size();
      l.next = text.size();
    } else {
      l.end = nl;
      l.next = nl + 1;
    }
    if (l.end > l.begin && text[l.end - 1] == '\r')
      --l.end;
    lines.push_back(l);
    pos = l.next;
  }
  return lines;
}

// Walks the header of a dict.yaml and finds the `file:` and `sheet:` lines
// inside the `excel:` block.
struct Found {
  int file_line = -1, sheet_line = -1;
  int excel_line = -1;           // the `excel:` line
  int last_excel_child = -1;     // last indented line of the block
  size_t child_indent = 2;
};

inline Found Locate(const std::string& text, const std::vector<Line>& lines) {
  Found f;
  bool in_excel = false;
  for (size_t i = 0; i < lines.size(); ++i) {
    std::string line = text.substr(lines[i].begin, lines[i].end - lines[i].begin);
    if (i == 0 && line.compare(0, 3, "\xEF\xBB\xBF") == 0)
      line = line.substr(3);
    std::string t = Trim(line);
    if (t == "...")
      break;
    if (t.empty() || t[0] == '#' || t == "---")
      continue;
    bool indented = line[0] == ' ' || line[0] == '\t';
    if (!indented) {
      in_excel = t.compare(0, 6, "excel:") == 0;
      if (in_excel)
        f.excel_line = static_cast<int>(i);
      continue;
    }
    if (!in_excel)
      continue;
    f.last_excel_child = static_cast<int>(i);
    size_t indent = line.find_first_not_of(" \t");
    f.child_indent = indent == std::string::npos ? 2 : indent;
    if (t.compare(0, 5, "file:") == 0 && f.file_line < 0)
      f.file_line = static_cast<int>(i);
    else if (t.compare(0, 6, "sheet:") == 0 && f.sheet_line < 0)
      f.sheet_line = static_cast<int>(i);
  }
  return f;
}

}  // namespace detail

// Reads file / sheet from the excel: block. False if there is no such block
// or no file: entry.
inline bool ParseExcelSettings(const std::string& text, ExcelSettings* out) {
  auto lines = detail::SplitLines(text);
  detail::Found f = detail::Locate(text, lines);
  if (f.excel_line < 0 || f.file_line < 0)
    return false;
  auto value_of = [&](int idx, size_t key_len) {
    std::string line =
        text.substr(lines[idx].begin, lines[idx].end - lines[idx].begin);
    std::string t = detail::Trim(line);
    return detail::Scalar(t.substr(key_len));
  };
  out->file = value_of(f.file_line, 5);
  out->sheet = f.sheet_line >= 0 ? value_of(f.sheet_line, 6) : std::string();
  return !out->file.empty();
}

// Returns `text` with the file: and sheet: values replaced (line breaks, BOM,
// comments and every other line are kept). False if the block cannot be found.
inline bool ReplaceExcelSettings(const std::string& text,
                                 const ExcelSettings& settings,
                                 std::string* result) {
  auto lines = detail::SplitLines(text);
  detail::Found f = detail::Locate(text, lines);
  if (f.excel_line < 0 || f.file_line < 0)
    return false;
  size_t first_nl = text.find('\n');
  std::string eol =
      (first_nl != std::string::npos && first_nl > 0 && text[first_nl - 1] == '\r')
          ? "\r\n"
          : "\n";
  std::string indent(f.child_indent, ' ');
  std::string out;
  for (size_t i = 0; i < lines.size(); ++i) {
    std::string raw = text.substr(lines[i].begin, lines[i].next - lines[i].begin);
    if (static_cast<int>(i) == f.file_line) {
      out += indent + "file: " + detail::Quote(settings.file) + eol;
    } else if (static_cast<int>(i) == f.sheet_line) {
      out += indent + "sheet: " + detail::Quote(settings.sheet) + eol;
    } else {
      out += raw;
    }
    if (static_cast<int>(i) == f.last_excel_child && f.sheet_line < 0) {
      // no sheet: line yet: add one after the last line of the block
      if (raw.empty() || raw.back() != '\n')
        out += eol;
      out += indent + "sheet: " + detail::Quote(settings.sheet) + eol;
    }
  }
  *result = out;
  return true;
}

// Entry count stored in the metadata at the start of a compiled
// build/*.table.bin ("Rime::Table/4.0"). False if the file is missing or not a
// table.
inline bool ReadTableEntryCount(const std::filesystem::path& table_file,
                                uint32_t* entries) {
  std::ifstream fin(table_file, std::ios::in | std::ios::binary);
  if (!fin)
    return false;
  unsigned char head[44] = {0};
  fin.read(reinterpret_cast<char*>(head), sizeof(head));
  if (fin.gcount() < static_cast<std::streamsize>(sizeof(head)))
    return false;
  if (std::string(reinterpret_cast<char*>(head), 11) != "Rime::Table")
    return false;
  // format[32], checksum, num_syllables, num_entries (little endian)
  *entries = static_cast<uint32_t>(head[40]) |
             (static_cast<uint32_t>(head[41]) << 8) |
             (static_cast<uint32_t>(head[42]) << 16) |
             (static_cast<uint32_t>(head[43]) << 24);
  return true;
}

// ".xlsx" / ".xlsm": supported. ".xls": not supported (old binary format).
enum class WorkbookKind { kSupported, kOldXls, kOther };

inline WorkbookKind ClassifyWorkbookName(std::string name) {
  size_t dot = name.rfind('.');
  std::string ext = dot == std::string::npos ? "" : name.substr(dot);
  for (auto& c : ext)
    if (c >= 'A' && c <= 'Z')
      c = static_cast<char>(c - 'A' + 'a');
  if (ext == ".xlsx" || ext == ".xlsm")
    return WorkbookKind::kSupported;
  if (ext == ".xls")
    return WorkbookKind::kOldXls;
  return WorkbookKind::kOther;
}

}  // namespace excel_edition
