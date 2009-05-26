// File: ZipArchive.cpp
// ZipArchive implementation file
//

#include "ZipArchive.h"
#include <time.h>
#include <sstream>
#include "splibint.h"

#include <string.h>

namespace splib {

ZipArchive::ZipArchive(const _TCHAR* pathname) {
    zf = zipOpen(pathname, APPEND_STATUS_CREATE);
    if (zf == 0) {
        std::basic_string<_TCHAR> msg;
        msg += _T("error opening [");
        msg += pathname;
        msg += _T("] for writing");
        throw IOException(msg.c_str());
    }
    entryOpened = false;
}

ZipArchive::~ZipArchive() {
    if (entryOpened) {
        closeEntry();
    }
    if (zf != 0) {
        close();
    }

}

void ZipArchive::openEntry(const char* entryName) {
    if (zf == 0) {
        throw IllegalStateException();
    }
    if (entryOpened) {
        throw IllegalStateException();
    }
    time_t timeValue = time(0);
#pragma warning (disable : 4996)
    struct tm* timeStruct = localtime(&timeValue);
#pragma warning (default : 4996)
    zip_fileinfo zi;
    zi.tmz_date.tm_sec = timeStruct->tm_sec;
    zi.tmz_date.tm_min = timeStruct->tm_min;
    zi.tmz_date.tm_hour = timeStruct->tm_hour;
    zi.tmz_date.tm_mday = timeStruct->tm_mday;
    zi.tmz_date.tm_mon = timeStruct->tm_mon;
    zi.tmz_date.tm_year = timeStruct->tm_year + 1900;
    zi.dosDate = 0;
    zi.internal_fa = 0;
    zi.external_fa = 0;
    int err = zipOpenNewFileInZip(zf, entryName, &zi,
        0, 0, 0, 0, 0, Z_DEFLATED, Z_DEFAULT_COMPRESSION);
    if (err != ZIP_OK) {
        throw IOException(_T("error opening zip entry"));
    }
    entryOpened = true;
}

void ZipArchive::write(const void* buffer, unsigned length) {
    if (zf == 0) {
        throw IllegalStateException();
    }
    if (!entryOpened) {
        throw IllegalStateException();
    }
    if (zipWriteInFileInZip(zf, buffer, length) < 0) {
        throw IOException(_T("error writing zip entry"));
    }
}

void ZipArchive::closeEntry() {
    if (zf == 0) {
        throw IllegalStateException();
    }
    if (!entryOpened) {
        throw IllegalStateException();
    }
    if (zipCloseFileInZip(zf) != ZIP_OK) {
        throw IOException(_T("error closing current zip entry"));
    }
    entryOpened = false;
}

void ZipArchive::close() {
    if (zf == 0) {
        throw IllegalStateException();
    }
    if (entryOpened) {
        closeEntry();
    }
    if (zipClose(zf, 0) != ZIP_OK) {
        throw IOException(_T("error closing zip archive"));
    }
    zf = 0;
}

ZipArchive& operator << (ZipArchive& ar, int val) {
    std::basic_stringstream<char> str;
    str << val;
    ar << str.str().c_str();
    return ar;
}

ZipArchive& operator << (ZipArchive& ar, long val) {
    std::basic_stringstream<char> str;
    str << val;
    ar << str.str().c_str();
    return ar;
}

ZipArchive& operator << (ZipArchive& ar, double val) {
    std::basic_stringstream<char> str;
    str.precision(17);
    str << val;
    ar << str.str().c_str();
    return ar;
}

ZipArchive& operator << (ZipArchive& ar, const char* val) {
    ar.write(val, (unsigned)::strlen(val));
    return ar;
}

}
