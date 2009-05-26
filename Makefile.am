lib_LTLIBRARIES = libsplib.la
libsplib_la_SOURCES = \
include/splib.h \
src/BasicExcel.cpp src/BasicExcel.h \
src/ByteArray.cpp src/ByteArray.h \
src/CellImpl.cpp src/CellImpl.h \
src/CellsImpl.cpp src/CellsImpl.h \
src/ColumnImpl.cpp src/ColumnImpl.h \
src/ColumnsImpl.cpp src/ColumnsImpl.h \
src/Date.cpp \
src/ExcelUtil.cpp src/ExcelUtil.h \
src/ExceptionImpl.cpp \
src/Formulas.cpp src/Formulas.h \
src/IllegalArgumentException.cpp \
src/IllegalStateException.cpp \
src/IndexedCollectionImpl.h \
src/IndexLimits.cpp src/IndexLimits.h \
src/ioapi.cpp src/ioapi.h \
src/IOException.cpp \
src/OdsWriter.cpp \
src/OdsWriterImpl.cpp src/OdsWriterImpl.h \
src/RowImpl.cpp src/RowImpl.h \
src/RowsImpl.cpp src/RowsImpl.h \
src/splib.cpp \
src/splibint.h \
src/SpreadsheetImpl.cpp \
src/Strings.cpp src/Strings.h \
src/TableImpl.cpp src/TableImpl.h \
src/Time.cpp \
src/ToUTF16.cpp src/ToUTF16.h \
src/ToUTF8.cpp src/ToUTF8.h \
src/Util.cpp src/Util.h \
src/XlsWriter.cpp \
src/XlsWriterImpl.cpp src/XlsWriterImpl.h \
src/XlsxWriter.cpp \
src/XlsxWriterImpl.cpp src/XlsxWriterImpl.h \
src/zip.cpp src/zip.h \
src/ZipArchive.cpp src/ZipArchive.h

libsplib_la_LDFLAGS = -version-info ${LIBRARY_VERSION_INFO} ${LTLIBICONV} -lz
libsplib_la_CPPFLAGS = -I${top_srcdir}/include

library_includedir = $(includedir)/splib
library_include_HEADERS = include/splib.h

pkgconfigdir = $(libdir)/pkgconfig
pkgconfig_DATA = ${PACKAGE_NAME}.pc

ACLOCAL_AMFLAGS = -I m4