cmake_minimum_required(VERSION 2.6)
project(SPREADSHEET)

find_package(ZLIB)
# Add iconv!
#find_package(Doxygen)
# Bring doxygen back!

set(CMAKE_CXX_FLAGS "-Wall -fpermissive")

add_subdirectory(src)
add_subdirectory(test)
