#include "xls2csv.h"

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <pybind11/stl_bind.h>

#include <algorithm>
#include <atomic>
#include <chrono>
#include <execution>
#include <filesystem>
#include <map>
#include <shared_mutex>
#include <sstream>
#include <string>
#include <vector>

#include <xls.h>



namespace fs = std::filesystem;
namespace py = pybind11;
using namespace py::literals;
using namespace std;
using namespace xls;

string tolower(string_view c) {
  string r;
  for (auto i : c) {
    r += tolower(i);
  }
  return r;
}

struct Profiler {
  using clock = chrono::steady_clock;
  clock::time_point begin = clock::now();
  const char* name;
  bool enable;

  Profiler(const char* n, bool enable) : name{n} {}
  ~Profiler() {
    if (!enable)
      return;
    auto end = clock::now();
    auto cost =
        chrono::duration_cast<chrono::milliseconds>(end - begin).count();
    printf("%s cost %.2fs.\n", name, cost / 1000.0f);
  }
};

struct Col {
  string name;

  enum Type {
    Comment,
    Int,
    Float,
    Str,
    Table,
  };
  Type type = Comment;

  Type parseType(string_view s) {
    auto l = tolower(s);
    if (l == "int")
      return Int;
    else if (l == "float")
      return Float;
    else if (l == "str")
      return Str;
    else if (l == "table")
      return Table;
    return Comment;
  }
};

using StrMap = std::map<string, string>;

struct Cache {
  StrMap files;
  std::shared_mutex mu;

  void addFile(string f, string data) {
    std::unique_lock l{mu};
    files[f] = std::move(data);
  }
};

bool isNumber(int id) {
  return id == XLS_RECORD_RK || id == XLS_RECORD_MULRK ||
         id == XLS_RECORD_NUMBER || id == XLS_RECORD_FORMULA;
}

tuple<double, bool> isStrDouble(const char* s) {
  if (!s)
    return {};

  char* ep;
  if (auto p = strtod(s, &ep); ep == s + strlen(s)) {
    return {p, true};
  }
  return {};
}

string numberToString(double p) {
  if (round(p) == p) {
    return to_string((int64_t)p);
  } else {
    return to_string(p);
  }
}

void parseXls(string fname,
              string_view out_dir,
              Cache& cache,
              const StrMap& ignoreXls,
              const StrMap& ignoreSheets) {
  if (strstr(fname.data(), ".xlsx") || !strstr(fname.data(), ".xls")) {
    return;
  }
  if (auto i = ignoreXls.find(fname); i != ignoreXls.end()) {
    return;
  }

  xls_error_t error = LIBXLS_OK;
  xlsWorkBook* wb = xls_open_file(fname.data(), "UTF-8", &error);
  if (wb == NULL) {
    printf("Error reading file: %s %s\n", xls_getError(error), fname.data());
    return;
  }

  for (int sheetIdx = 0; sheetIdx < wb->sheets.count; sheetIdx++) {
    xlsWorkSheet* work_sheet = xls_getWorkSheet(wb, sheetIdx);
    string_view sheet_name = wb->sheets.sheet[sheetIdx].name;

    auto& base = fname.substr(fname.rfind('\\') + 1);
    string out_sheet =
        tolower(base.substr(0, base.rfind('.'))) + "_" + tolower(sheet_name);
    string csv_name = out_sheet + ".csv";

    if (auto it = ignoreSheets.find(out_sheet); it != ignoreSheets.end()) {
      continue;
    }

    map<int, Col> header;
    stringstream csv;
    error = xls_parseWorkSheet(work_sheet);
    if (error != LIBXLS_OK) {
      printf("Error reading sheet: %s %s %s\n", xls_getError(error),
             fname.data(), sheet_name.data());
      return;
    }
    int realColCnt = 0;
    int rowIdxEnd = work_sheet->rows.lastrow, colEnd = work_sheet->rows.lastcol;

    for (int rowIdx = 0; rowIdx <= rowIdxEnd; rowIdx++) {
      xlsRow* row = xls_row(work_sheet, rowIdx);
      auto cells = row->cells.cell;

      // header
      if (rowIdx == 0) {
        for (int colIdx = 0; colIdx <= colEnd; colIdx++, realColCnt++) {
          auto c = cells[colIdx];
          auto& h = header[colIdx];

          if (c.id == XLS_RECORD_BLANK) {
            // end fo columns
            if (all_of(cells + colIdx, cells + colEnd,
                       [](auto& c) { return c.id == XLS_RECORD_BLANK; })) {
              realColCnt = colIdx;
              break;
            }
            continue;
          }

          if (isNumber(c.id)) {
            // invalid header, ignore
            printf("invalid head: %s\n", out_sheet.c_str());
            return;
          }

          h.name = c.str;
          if (auto p = h.name.find_first_of("_"); p != string::npos) {
            if (auto t = h.parseType(h.name.substr(0, p)); t != Col::Comment) {
              h.type = t;
              h.name = h.name.substr(p + 1);
            }
          }
        }

        for (int colIdx = 0; colIdx < realColCnt; colIdx++) {
          if (header[colIdx].type == Col::Comment)
            continue;
          if (colIdx != 0)
            csv << ",";
          csv << header[colIdx].name;
        }
        csv << "\n";
        continue;
      }

      // empty line, end of data
      if (std::all_of(cells, cells + realColCnt,
                      [](auto& c) { return c.id == XLS_RECORD_BLANK; })) {
        break;
      }

      string csvrow;
      csvrow.reserve(1024);
      for (int colIdx = 0; colIdx < realColCnt; colIdx++) {
        auto& colInfo = header[colIdx];
        if (colInfo.type == Col::Comment) {
          continue;
        }

        if (colIdx != 0) {
          csvrow += ',';
        }

        xlsCell* cell = &cells[colIdx];
        switch (colInfo.type) {
          case Col::Int: {
            csvrow += to_string(lround(cell->d));
          } break;

          case Col::Float: {
            if (cell->d == 0)
              csvrow += "0.0";
            else
              csvrow += to_string(cell->d);
          } break;

          case Col::Str: {
            csvrow += "\"";
            if (cell->str) {
              if (auto [p, ok] = isStrDouble(cell->str); ok) {
                csvrow += numberToString(p);
              } else {
                csvrow += cell->str;
              }
            } else if (isNumber(cell->id)) {
              csvrow += numberToString(cell->d);
            }
            csvrow += "\"";
          } break;

          case Col::Table: {
            csvrow += "\"";
            if (cell->str)
              csvrow += cell->str;
            csvrow += "\"";
          } break;
        }
      }

      csv << csvrow << '\n';
    }

    xls_close_WS(work_sheet);

    if (!out_dir.empty()) {
      stringstream ss;
      ss << out_dir << "\\" << csv_name;
      if (auto fp = fopen(ss.str().c_str(), "w")) {
        auto& data = csv.str();
        fwrite(data.c_str(), data.length(), 1, fp);
        fclose(fp);
      }
    }

    cache.addFile(csv_name, std::move(csv.str()));
  }

  xls_close_WB(wb);
}

auto vec2map(const vector<string>& v) {
  StrMap m;
  for (auto i : v) {
    m[i] = i;
  }
  return m;
}

StrMap loadXls(string path,
               string out_dir,
               const vector<string>& ignoreXls,
               const vector<string>& ignoreSheetName,
               bool profile = true) {
  Cache data;
  StrMap ignoreSheetMap = vec2map(ignoreSheetName);
  StrMap ignoreXlsMap = vec2map(ignoreXls);

  Profiler loadXmlProf("load xls", profile);

#ifdef _DEBUG
  parseXls(R"(C:\Samo\Trunk\data\GameDatas\datas\season.xls)", {}, data, {},
           {});
#endif

  auto files = fs::directory_iterator(path);
  std::vector<fs::directory_entry> file_names{
      fs::begin(files),
      fs::end(files),
  };

  if (!out_dir.empty() && !fs::is_directory(out_dir))
    fs::create_directory(out_dir);

#ifdef _DEBUG
  auto policy = std::execution::seq;
#else
  auto policy = std::execution::par;
#endif

  std::for_each(policy, file_names.begin(), file_names.end(),
                [&](const auto& entry) {
                  parseXls(entry.path().string(), out_dir, data, ignoreXlsMap,
                           ignoreSheetMap);
                });
  if (profile)
    printf("total xls files: %d\n", (int)data.files.size());
  return std::move(data.files);
}

PYBIND11_MAKE_OPAQUE(StrMap);

PYBIND11_MODULE(xls2csv, m) {
  py::bind_map<StrMap>(m, "StrStrMap");

  m.def("load_xls_dir", loadXls, "path"_a, "out_dir"_a = "",
        "ignore_xls"_a = vector<string>{}, "ignore_sheets"_a = vector<string>{},
        "profile"_a = true);
}

#ifdef TEST
int main() {
  loadXls(R"(C:\Samo\Trunk\data\GameDatas\datas)", "tmp", {}, {"fte_readme"});
  return 0;
}
#endif