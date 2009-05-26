// File: Formulas.cpp
// Formulas implementation file
//

#include "Formulas.h"
#include "Util.h"
#include "splibint.h"

#include <string.h>

namespace splib {

bool Formulas::parse(const _TCHAR* formula, int& c1, int& r1, 
                     int& c2, int& r2) {
    int a, b, c, d;
    if (!advance(formula, _T("SUM("))) {
        return false;
    }
    if (!Util::parseEmbeddedLocation(formula, a, b)) {
        return false;
    }
    if (!advance(formula, _T(":"))) {
        return false;
    }
    if (!Util::parseEmbeddedLocation(formula, c, d)) {
        return false;
    }
    if (!advance(formula, _T(")"))) {
        return false;
    }
    if (*formula) {
        return false;
    }
    c1 = a;
    r1 = b;
    c2 = c;
    r2 = d;
    return true;
}

bool Formulas::advance(const _TCHAR*& p, const _TCHAR* sequence) {
    size_t len = ::_tcslen(sequence);
    if (::_tcsncmp(p, sequence, len) == 0) {
        p += len;
        return true;
    } else {
        return false;
    }
}

}
