// File: Util.cpp
// Util implementation file
//

#include "Util.h"
#include <sstream>
#include "splibint.h"

namespace splib {

bool Util::parseLocation(const _TCHAR* str, int& col, int& row) {
    int c, r;
    if (!parseEmbeddedLocation(str, c, r)) {
        return false;
    }
    if (*str) {
        return false;
    }
    col = c;
    row = r;
    return true;
}

bool Util::parseEmbeddedLocation(const _TCHAR*& str, int& col, int& row) {
    // parse letter part (determine its length first)
    const _TCHAR* s = str;
    int len = 0;
    while(*s >= 'A' && *s <= 'Z') {
        s++;
        len++;
    }
    if (len < 1) {
        return false;
    }
    s = str;
    // now actual parsing
    int c = 0;
    int BASE = 'Z' - 'A' + 1;
    while(*s >= 'A' && *s <= 'Z') {
        c = c * BASE + (*s - 'A');
        // the most significant symbol should be +1 if the length
        // is greater than 1. That's why that initial length
        // calculation was necessary
        if (s == str && len > 1) {
            c++;
        }
        s++;
    }
    // parse number part
    int r = 0;
    bool specified = false;
    BASE = 10;
    while(*s >= '0' && *s <= '9') {
        r = r * BASE + (*s - '0');
        specified = true;
        s++;
    }
    if (!specified) {
        return false;
    }
    // 0 is not allowed
    if (r < 1) {
        return false;
    }
    r--;

    str = s;
    col = c;
    row = r;
    return true;
}

bool Util::parseColumn(const _TCHAR* str, int& col) {
    // determine length first
    const _TCHAR* s = str;
    int len = 0;
    while(*s >= 'A' && *s <= 'Z') {
        s++;
        len++;
    }
    if (len < 1) {
        return false;
    }
    s = str;
    // now actual parsing
    int c = 0;
    int BASE = 'Z' - 'A' + 1;
    while(*s >= 'A' && *s <= 'Z') {
        c = c * BASE + (*s - 'A');
        // the most significant symbol should be +1 if the length
        // is greater than 1. That's why that initial length
        // calculation was necessary
        if (s == str && len > 1) {
            c++;
        }
        s++;
    }
    if (*s) {
        return false;
    }

    col = c;
    return true;
}

std::basic_string<char> Util::buildLocation(int col, int row) {
    std::basic_string<char> column;
    const int BASE = 'Z' - 'A' + 1;
    do {
        char c = 'A' + (char)(col % BASE);
        if (column.length() > 0) {
            c--;
        }
        column.insert(column.begin(), c);
        col = col / BASE;
    } while (col > 0);
    std::basic_stringstream<char> res;
    res << column << row + 1;
    return res.str();
}

}
