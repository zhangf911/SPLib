// File: ToUTF8.cpp
// ToUTF8 implementation file
//

#include "ToUTF8.h"
#ifdef WIN32
#include <windows.h>
#endif
#include "splibint.h"

#include <string.h>

namespace splib {

ToUTF8::ToUTF8(const _TCHAR* s) {
#ifdef WIN32
#ifdef _UNICODE
    // convert from wide char to UTF8 char
    int len = ::WideCharToMultiByte(CP_UTF8, 0, s, -1, 0, 0, 0, 0);
    buffer = new char[len];
    ::WideCharToMultiByte(CP_UTF8, 0, s, -1, (char*)buffer, len, 0, 0);
#else // !_UNICODE
    // convert from CP_ACP char to wide char
    // and then from wide char to UTF8 char
    int tmplen = ::MultiByteToWideChar(CP_ACP, 0, s, -1, 0, 0);
    wchar_t* tmp = new wchar_t[tmplen];
    ::MultiByteToWideChar(CP_ACP, 0, s, -1, tmp, tmplen);
    int len = WideCharToMultiByte(CP_UTF8, 0, tmp, -1, 0, 0, 0, 0);
    buffer = new char[len];
    ::WideCharToMultiByte(CP_UTF8, 0, tmp, -1, (char*)buffer, len, 0, 0);
    delete [] tmp;
#endif // _UNICODE
    ownBuffer = true;
#else // !WIN32
    // on other systems the input string is assumed to be UTF8 already
    buffer = s;
    ownBuffer = false;
#endif // WIN32
}

ToUTF8::~ToUTF8() {
    if (ownBuffer) {
        delete [] buffer;
    }
}

const char* ToUTF8::get() const {
    return buffer;
}

int ToUTF8::length() const {
    return (int)::strlen(buffer);
}

}
