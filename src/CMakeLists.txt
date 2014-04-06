IF( WIN32 )
	SET( CMAKE_CXX_FLAGS "-D NDEBUG -D WIN32 -D _WINDOWS -D _USRDLL -D SPLIB_EXPORTS -D UNICODE -D _UNICODE")
ENDIF( WIN32 )

add_library(spreadsheet SHARED splib.h 
BasicExcel.cpp BasicExcel.h 
ByteArray.cpp ByteArray.h 
CellImpl.cpp CellImpl.h 
CellsImpl.cpp CellsImpl.h 
ColumnImpl.cpp ColumnImpl.h 
ColumnsImpl.cpp ColumnsImpl.h 
Date.cpp 
ExcelUtil.cpp ExcelUtil.h 
ExceptionImpl.cpp 
Formulas.cpp Formulas.h 
IllegalArgumentException.cpp 
IllegalStateException.cpp 
IndexedCollectionImpl.h 
IndexLimits.cpp IndexLimits.h 
ioapi.cpp ioapi.h 
IOException.cpp 
OdsWriter.cpp 
OdsWriterImpl.cpp OdsWriterImpl.h 
RowImpl.cpp RowImpl.h 
RowsImpl.cpp RowsImpl.h 
splib.cpp 
splibint.h 
SpreadsheetImpl.cpp 
Strings.cpp Strings.h 
TableImpl.cpp TableImpl.h 
Time.cpp 
ToUTF16.cpp ToUTF16.h 
ToUTF8.cpp ToUTF8.h 
Util.cpp Util.h 
XlsWriter.cpp 
XlsWriterImpl.cpp XlsWriterImpl.h 
XlsxWriter.cpp 
XlsxWriterImpl.cpp XlsxWriterImpl.h 
zip.cpp zip.h 
ZipArchive.cpp ZipArchive.h)

# target_link_libraries(spreadsheet zlib libiconv.dll)
target_link_libraries(spreadsheet z c)
