find_package(ZLIB REQUIRED)

IF( WIN32 )
	SET( CMAKE_CXX_FLAGS "-D WIN32 -D _WINDOWS -D UNICODE -D _UNICODE")
ENDIF( WIN32 )

add_executable(test test.cpp)

include_directories(../src)
target_link_libraries(test spreadsheet ${ZLIB_LIBRARIES})