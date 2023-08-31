cmake_minimum_required(VERSION 3.16)
project ("xls2csv")

set (CMAKE_CXX_STANDARD 17)
set (CMAKE_CXX_STANDARD_REQUIRED True)


find_package(Python 3.11 COMPONENTS Interpreter Development REQUIRED)
add_subdirectory(pybind11)


add_library (xls2csv SHARED
	"xls2csv.cpp" 
	"xls2csv.h" 
	"libxls/src/endian.c"
	"libxls/src/locale.c"
	"libxls/src/ole.c"
	"libxls/src/xls.c"	 
	"libxls/src/xlstool.c"	 
	"libxls/include/config.h"
)
target_compile_definitions(xls2csv PUBLIC -DXLS2CSV_PATCH)
target_include_directories(xls2csv PUBLIC "libxls/include")
target_link_libraries (xls2csv
	PUBLIC Python::Python
	PUBLIC pybind11::pybind11
)


add_executable (xls2csv_test 
	"xls2csv.cpp" 
	"xls2csv.h" 
	"libxls/src/endian.c"
	"libxls/src/locale.c"
	"libxls/src/ole.c"
	"libxls/src/xls.c"	 
	"libxls/src/xlstool.c"	 
	"libxls/include/config.h"
)
target_compile_definitions(xls2csv_test PUBLIC -DTEST -DXLS2CSV_PATCH)
target_include_directories(xls2csv_test PUBLIC "libxls/include")
target_link_libraries (xls2csv_test
	PUBLIC Python::Python
	PUBLIC pybind11::pybind11
)