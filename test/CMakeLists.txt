find_package(ZLIB REQUIRED)

add_executable(test test.cpp)

include_directories(../src)
target_link_libraries(test spreadsheet ${ZLIB_LIBRARIES})