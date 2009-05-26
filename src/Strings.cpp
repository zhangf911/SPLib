// File: Strings.cpp
// Strings implementation file
//

#include "Strings.h"
#include "splibint.h"

#include <string.h>

namespace splib {

_tstring Strings::xmlize(const _TCHAR* str) {
    _tstring res = str;
    res = replace(res.c_str(), _T("&"), _T("&amp;"));
    res = replace(res.c_str(), _T("<"), _T("&lt;"));
    res = replace(res.c_str(), _T(">"), _T("&gt;"));
    res = replace(res.c_str(), _T("\""), _T("&quot;"));
    return res;
}

_tstring Strings::replace(const _TCHAR* str, const _TCHAR* substring,
                          const _TCHAR* replacement) {
    _tstring res = str;
    _tstring::size_type pos = 0;
    size_t substringlen = ::_tcslen(substring);
    size_t replacementlen = ::_tcslen(replacement);
#pragma warning (disable : 4127)
    while (true) {
#pragma warning (default : 4127)
        pos = res.find(substring, pos);
        if (pos == _tstring::npos) {
            break;
        }
        res.replace(pos, substringlen, replacement);
        pos += replacementlen;
    }
    return res;
}

}
