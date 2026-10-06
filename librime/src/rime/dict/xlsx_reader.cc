//
// Copyright RIME Developers
// Distributed under the BSD License
//
// Minimal reader for Excel .xlsx workbooks (Office Open XML).
// Only cell values are read; formatting, formulas and other parts of the
// workbook are ignored. Works with files saved by Microsoft Excel and WPS.
//
#include <algorithm>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <sstream>
#include <rime/dict/xlsx_reader.h>
#include "miniz/miniz.h"

#ifdef _WIN32
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#endif

namespace rime {

namespace {

constexpr size_t npos = string::npos;

struct XmlTag {
  size_t start = npos;   // position of '<'
  size_t end = npos;     // position just after '>'
  string local_name;     // element name without namespace prefix
  bool closing = false;  // </name>
  bool self_closing = false;  // <name ... />
  string attrs;          // raw attribute text
};

// Finds the next element tag at or after pos. Skips declarations,
// processing instructions and comments.
bool NextTag(const string& xml, size_t pos, XmlTag* tag) {
  while (true) {
    pos = xml.find('<', pos);
    if (pos == npos)
      return false;
    if (xml.compare(pos, 4, "<!--") == 0) {
      size_t e = xml.find("-->", pos + 4);
      if (e == npos)
        return false;
      pos = e + 3;
      continue;
    }
    if (xml.compare(pos, 9, "<![CDATA[") == 0) {
      size_t e = xml.find("]]>", pos + 9);
      if (e == npos)
        return false;
      pos = e + 3;
      continue;
    }
    size_t gt = xml.find('>', pos);
    if (gt == npos)
      return false;
    if (xml[pos + 1] == '?' || xml[pos + 1] == '!') {
      pos = gt + 1;
      continue;
    }
    tag->start = pos;
    tag->end = gt + 1;
    size_t p = pos + 1;
    tag->closing = (xml[p] == '/');
    if (tag->closing)
      ++p;
    size_t name_end = p;
    while (name_end < gt && xml[name_end] != ' ' && xml[name_end] != '\t' &&
           xml[name_end] != '\r' && xml[name_end] != '\n' &&
           xml[name_end] != '/' && xml[name_end] != '>')
      ++name_end;
    string name = xml.substr(p, name_end - p);
    size_t colon = name.find(':');
    tag->local_name = colon == npos ? name : name.substr(colon + 1);
    tag->self_closing = !tag->closing && gt > pos && xml[gt - 1] == '/';
    size_t attrs_end = tag->self_closing ? gt - 1 : gt;
    tag->attrs = name_end < attrs_end
                     ? xml.substr(name_end, attrs_end - name_end)
                     : string();
    return true;
  }
}

void AppendUtf8(uint32_t cp, string* out) {
  if (cp < 0x80) {
    out->push_back(static_cast<char>(cp));
  } else if (cp < 0x800) {
    out->push_back(static_cast<char>(0xC0 | (cp >> 6)));
    out->push_back(static_cast<char>(0x80 | (cp & 0x3F)));
  } else if (cp < 0x10000) {
    out->push_back(static_cast<char>(0xE0 | (cp >> 12)));
    out->push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
    out->push_back(static_cast<char>(0x80 | (cp & 0x3F)));
  } else if (cp < 0x110000) {
    out->push_back(static_cast<char>(0xF0 | (cp >> 18)));
    out->push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3F)));
    out->push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
    out->push_back(static_cast<char>(0x80 | (cp & 0x3F)));
  }
}

bool IsHex(char c) {
  return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') ||
         (c >= 'A' && c <= 'F');
}

// Decodes XML entities and Excel's _xHHHH_ escapes.
string DecodeText(const string& raw) {
  string out;
  out.reserve(raw.size());
  for (size_t i = 0; i < raw.size(); ++i) {
    char c = raw[i];
    if (c == '&') {
      size_t semi = raw.find(';', i);
      if (semi != npos && semi - i <= 10) {
        string ent = raw.substr(i + 1, semi - i - 1);
        bool ok = true;
        if (ent == "lt")
          out.push_back('<');
        else if (ent == "gt")
          out.push_back('>');
        else if (ent == "amp")
          out.push_back('&');
        else if (ent == "quot")
          out.push_back('"');
        else if (ent == "apos")
          out.push_back('\'');
        else if (ent.size() > 1 && ent[0] == '#') {
          try {
            uint32_t cp = (ent[1] == 'x' || ent[1] == 'X')
                              ? std::stoul(ent.substr(2), nullptr, 16)
                              : std::stoul(ent.substr(1), nullptr, 10);
            AppendUtf8(cp, &out);
          } catch (...) {
            ok = false;
          }
        } else {
          ok = false;
        }
        if (ok) {
          i = semi;
          continue;
        }
      }
    } else if (c == '_' && i + 6 < raw.size() && raw[i + 1] == 'x' &&
               IsHex(raw[i + 2]) && IsHex(raw[i + 3]) && IsHex(raw[i + 4]) &&
               IsHex(raw[i + 5]) && raw[i + 6] == '_') {
      AppendUtf8(std::stoul(raw.substr(i + 2, 4), nullptr, 16), &out);
      i += 6;
      continue;
    }
    out.push_back(c);
  }
  return out;
}

// Returns the value of the attribute with the given local name,
// ignoring namespace declarations.
string GetAttr(const string& attrs, const string& local_name) {
  size_t i = 0;
  while (i < attrs.size()) {
    while (i < attrs.size() && (attrs[i] == ' ' || attrs[i] == '\t' ||
                                attrs[i] == '\r' || attrs[i] == '\n'))
      ++i;
    size_t eq = attrs.find('=', i);
    if (eq == npos)
      break;
    string name = attrs.substr(i, eq - i);
    while (!name.empty() && (name.back() == ' ' || name.back() == '\t'))
      name.pop_back();
    size_t q = eq + 1;
    while (q < attrs.size() && attrs[q] == ' ')
      ++q;
    if (q >= attrs.size())
      break;
    char quote = attrs[q];
    if (quote != '"' && quote != '\'')
      break;
    size_t qe = attrs.find(quote, q + 1);
    if (qe == npos)
      break;
    if (name.compare(0, 5, "xmlns") != 0) {
      size_t colon = name.find(':');
      string local = colon == npos ? name : name.substr(colon + 1);
      if (local == local_name)
        return DecodeText(attrs.substr(q + 1, qe - q - 1));
    }
    i = qe + 1;
  }
  return string();
}

// Text between the end of an opening tag and the next '<'.
string TextAfter(const string& xml, const XmlTag& tag) {
  if (tag.self_closing)
    return string();
  size_t lt = xml.find('<', tag.end);
  if (lt == npos)
    return string();
  return xml.substr(tag.end, lt - tag.end);
}

// "AB12" -> 27 (0-based column index); returns -1 if no letters.
int ColumnIndex(const string& ref) {
  int col = 0;
  size_t i = 0;
  for (; i < ref.size(); ++i) {
    char c = ref[i];
    if (c >= 'A' && c <= 'Z')
      col = col * 26 + (c - 'A' + 1);
    else if (c >= 'a' && c <= 'z')
      col = col * 26 + (c - 'a' + 1);
    else
      break;
  }
  return i == 0 ? -1 : col - 1;
}

// Resolves a relationship target relative to the "xl/" folder.
string ResolveTarget(const string& target) {
  string full = (!target.empty() && target[0] == '/') ? target.substr(1)
                                                       : "xl/" + target;
  vector<string> parts;
  std::stringstream ss(full);
  string seg;
  while (std::getline(ss, seg, '/')) {
    if (seg.empty() || seg == ".")
      continue;
    if (seg == "..") {
      if (!parts.empty())
        parts.pop_back();
      continue;
    }
    parts.push_back(seg);
  }
  string result;
  for (const auto& p : parts) {
    if (!result.empty())
      result += '/';
    result += p;
  }
  return result;
}

}  // namespace

bool XlsxReader::ReadFileBytes(const path& file_path, string* content) {
#ifdef _WIN32
  // Share everything so that a workbook kept open in Excel can still be read
  // and Excel is never blocked from saving.
  HANDLE h = CreateFileW(file_path.c_str(), GENERIC_READ,
                         FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                         NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
  if (h == INVALID_HANDLE_VALUE)
    return false;
  LARGE_INTEGER size;
  if (!GetFileSizeEx(h, &size) || size.QuadPart > 0x7FFFFFFF) {
    CloseHandle(h);
    return false;
  }
  content->resize(static_cast<size_t>(size.QuadPart));
  size_t done = 0;
  while (done < content->size()) {
    DWORD chunk = 0;
    DWORD want = static_cast<DWORD>(
        std::min<size_t>(content->size() - done, 1 << 24));
    if (!ReadFile(h, &(*content)[done], want, &chunk, NULL) || chunk == 0)
      break;
    done += chunk;
  }
  CloseHandle(h);
  content->resize(done);
  return true;
#else
  std::ifstream fin(file_path.c_str(), std::ios::in | std::ios::binary);
  if (!fin)
    return false;
  std::ostringstream buffer;
  buffer << fin.rdbuf();
  *content = buffer.str();
  return true;
#endif
}

bool XlsxReader::Open(const path& file_path) {
  data_.clear();
  shared_strings_.clear();
  error_.clear();
  if (!ReadFileBytes(file_path, &data_)) {
    error_ = "cannot open file: " + file_path.u8string();
    return false;
  }
  if (data_.size() < 4 || data_.compare(0, 2, "PK") != 0) {
    error_ = "not an .xlsx file (zip signature missing): " +
             file_path.u8string();
    return false;
  }
  return true;
}

bool XlsxReader::ExtractEntry(const string& entry_name,
                              string* content,
                              bool required) {
  mz_zip_archive zip;
  std::memset(&zip, 0, sizeof(zip));
  if (!mz_zip_reader_init_mem(&zip, data_.data(), data_.size(), 0)) {
    error_ = "broken .xlsx file (cannot read zip directory)";
    return false;
  }
  int index = mz_zip_reader_locate_file(&zip, entry_name.c_str(), nullptr, 0);
  if (index < 0) {
    mz_zip_reader_end(&zip);
    if (required)
      error_ = "missing part in .xlsx file: " + entry_name;
    return false;
  }
  size_t size = 0;
  void* p = mz_zip_reader_extract_to_heap(&zip, index, &size, 0);
  if (!p) {
    mz_zip_reader_end(&zip);
    error_ = "failed to decompress part: " + entry_name;
    return false;
  }
  content->assign(static_cast<const char*>(p), size);
  mz_free(p);
  mz_zip_reader_end(&zip);
  return true;
}

bool XlsxReader::LoadSharedStrings() {
  shared_strings_.clear();
  // locate sharedStrings part via workbook relationships
  string rels, part = "xl/sharedStrings.xml";
  if (ExtractEntry("xl/_rels/workbook.xml.rels", &rels, false)) {
    XmlTag tag;
    size_t pos = 0;
    while (NextTag(rels, pos, &tag)) {
      pos = tag.end;
      if (!tag.closing && tag.local_name == "Relationship") {
        string type = GetAttr(tag.attrs, "Type");
        if (type.size() >= 14 &&
            type.compare(type.size() - 14, 14, "/sharedStrings") == 0) {
          part = ResolveTarget(GetAttr(tag.attrs, "Target"));
        }
      }
    }
  }
  string xml;
  if (!ExtractEntry(part, &xml, false)) {
    return error_.empty();  // a workbook without text has no shared strings
  }
  XmlTag tag;
  size_t pos = 0;
  bool in_si = false;
  int skip_depth = 0;  // inside <rPh> (phonetic hints), ignored
  string current;
  while (NextTag(xml, pos, &tag)) {
    pos = tag.end;
    const string& n = tag.local_name;
    if (n == "si") {
      if (tag.closing) {
        shared_strings_.push_back(current);
        in_si = false;
      } else if (tag.self_closing) {
        shared_strings_.push_back(string());
      } else {
        in_si = true;
        current.clear();
      }
    } else if (n == "rPh" || n == "phoneticPr") {
      if (!tag.self_closing)
        skip_depth += tag.closing ? -1 : 1;
    } else if (n == "t" && in_si && !tag.closing && skip_depth == 0) {
      current += DecodeText(TextAfter(xml, tag));
    }
  }
  return true;
}

bool XlsxReader::ReadSheet(const string& sheet_name, vector<XlsxRow>* rows) {
  rows->clear();
  error_.clear();
  // 1. find the sheet's relationship id in workbook.xml
  string workbook;
  if (!ExtractEntry("xl/workbook.xml", &workbook, true))
    return false;
  string rel_id, available;
  {
    XmlTag tag;
    size_t pos = 0;
    while (NextTag(workbook, pos, &tag)) {
      pos = tag.end;
      if (tag.closing || tag.local_name != "sheet")
        continue;
      string name = GetAttr(tag.attrs, "name");
      if (!available.empty())
        available += ", ";
      available += name;
      if (rel_id.empty() && (sheet_name.empty() || name == sheet_name))
        rel_id = GetAttr(tag.attrs, "id");
    }
  }
  if (rel_id.empty()) {
    error_ = "sheet '" + sheet_name + "' not found; sheets in workbook: " +
             available;
    return false;
  }
  // 2. map relationship id to the worksheet part
  string rels;
  if (!ExtractEntry("xl/_rels/workbook.xml.rels", &rels, true))
    return false;
  string sheet_part;
  {
    XmlTag tag;
    size_t pos = 0;
    while (NextTag(rels, pos, &tag)) {
      pos = tag.end;
      if (!tag.closing && tag.local_name == "Relationship" &&
          GetAttr(tag.attrs, "Id") == rel_id) {
        sheet_part = ResolveTarget(GetAttr(tag.attrs, "Target"));
        break;
      }
    }
  }
  if (sheet_part.empty()) {
    error_ = "cannot locate worksheet part for relationship " + rel_id;
    return false;
  }
  if (!LoadSharedStrings())
    return false;
  string xml;
  if (!ExtractEntry(sheet_part, &xml, true))
    return false;
  // 3. walk rows and cells
  XmlTag tag;
  size_t pos = 0;
  bool in_row = false, in_cell = false, in_inline = false;
  int skip_depth = 0;
  XlsxRow row;
  int last_row_number = 0;
  int next_col = 0;
  int cell_col = 0;
  string cell_type, cell_value, inline_text;
  bool has_value = false;
  while (NextTag(xml, pos, &tag)) {
    pos = tag.end;
    const string& n = tag.local_name;
    if (n == "row") {
      if (tag.closing) {
        if (in_row)
          rows->push_back(std::move(row));
        in_row = false;
        continue;
      }
      row = XlsxRow();
      string r = GetAttr(tag.attrs, "r");
      row.row_number = r.empty() ? last_row_number + 1 : std::atoi(r.c_str());
      last_row_number = row.row_number;
      next_col = 0;
      in_row = !tag.self_closing;
      if (tag.self_closing)
        rows->push_back(std::move(row));
    } else if (n == "c" && in_row) {
      if (tag.closing) {
        if (!in_cell)
          continue;
        in_cell = false;
        string text;
        if (cell_type == "s") {
          if (has_value) {
            long idx = std::strtol(cell_value.c_str(), nullptr, 10);
            if (idx >= 0 && idx < static_cast<long>(shared_strings_.size()))
              text = shared_strings_[idx];
          }
        } else if (cell_type == "inlineStr") {
          text = inline_text;
        } else if (cell_type == "b") {
          text = cell_value == "1" ? "TRUE" : "FALSE";
        } else {
          text = DecodeText(cell_value);
        }
        if (static_cast<int>(row.cells.size()) <= cell_col)
          row.cells.resize(cell_col + 1);
        row.cells[cell_col] = std::move(text);
        continue;
      }
      string r = GetAttr(tag.attrs, "r");
      int col = r.empty() ? -1 : ColumnIndex(r);
      cell_col = col >= 0 ? col : next_col;
      next_col = cell_col + 1;
      cell_type = GetAttr(tag.attrs, "t");
      cell_value.clear();
      inline_text.clear();
      has_value = false;
      in_cell = !tag.self_closing;
    } else if (n == "v" && in_cell && !tag.closing) {
      cell_value = TextAfter(xml, tag);
      has_value = true;
    } else if (n == "is" && in_cell) {
      in_inline = !tag.closing && !tag.self_closing;
    } else if ((n == "rPh" || n == "phoneticPr") && in_cell) {
      if (!tag.self_closing)
        skip_depth += tag.closing ? -1 : 1;
    } else if (n == "t" && in_inline && !tag.closing && skip_depth == 0) {
      inline_text += DecodeText(TextAfter(xml, tag));
    }
  }
  return true;
}

}  // namespace rime
