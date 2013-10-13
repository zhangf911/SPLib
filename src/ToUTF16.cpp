// File: ToUTF16.cpp
// ToUTF16 implementation file
//

#include "ToUTF16.h"
#ifdef WIN32
# include <windows.h>
#else
# include <iconv.h>
// TODO use cmake to tell me if the second parameter must be const or not
// # ifndef ICONV_CONST
// #   define ICONV_CONST const
// # endif
# include <errno.h>
# include "ByteArray.h"
#endif
#include "splibint.h"

#include <string.h>

namespace splib {

ToUTF16::ToUTF16(const _TCHAR* s) {
#ifdef WIN32
#ifdef _UNICODE
    // no conversion necessary
    buffer = (const unsigned char*)s;
    len = (int)::_tcslen(s);
    ownBuffer = false;
#else // !_UNICODE
    // convert from char to wide char
    len = ::MultiByteToWideChar(CP_ACP, 0, s, -1, 0, 0);
    buffer = new unsigned char[len * 2];
    ::MultiByteToWideChar(CP_ACP, 0, s, -1, (WCHAR*)buffer, (int)len);
    len--;
    ownBuffer = true;
#endif
#else // !WIN32
    // use iconv for current locale to UTF16 conversion
    iconv_t cd = ::iconv_open("UTF-16LE", "");
    char* inbuf = (char*)s;
    size_t inbytesleft = ::strlen(s);
    const int outbufsize = 64;
    char outbufstg[outbufsize];
    char* outbuf = outbufstg;
    size_t outbytesleft = outbufsize;
    ByteArray byteArray;
    while (true) {
        size_t res = ::iconv(cd, &inbuf, &inbytesleft, &outbuf, &outbytesleft);
        if (res != -1) {
            size_t len = outbufsize - outbytesleft;
            byteArray.write((unsigned char*)outbufstg, len);
            break;
        } else if (res == -1 && errno == E2BIG) {
            size_t len = outbufsize - outbytesleft;
            byteArray.write((unsigned char*)outbufstg, len);
            outbuf = outbufstg;
            outbytesleft = outbufsize;
        } else {
            break;
        }
    }
    ::iconv_close(cd);
    // correction for BOM
    unsigned char bom[2] = {0xFF, 0xFE};
    if (::memcmp(byteArray.data(), bom, 2) == 0) {
        len = byteArray.size() - 2; // first 2 bytes are BOM
        buffer = new unsigned char[len];
        ::memcpy((void*)buffer, byteArray.data() + 2, len);
    } else {
        len = byteArray.size();
        buffer = new unsigned char[len];
        ::memcpy((void*)buffer, byteArray.data(), len);
    }
    len = len / 2;
    ownBuffer = true;
#endif // WIN32
}

ToUTF16::~ToUTF16() {
    if (ownBuffer) {
        delete [] buffer;
    }
}

const unsigned char* ToUTF16::get() const {
    return buffer;
}

int ToUTF16::length() const {
    return len;
}

}
