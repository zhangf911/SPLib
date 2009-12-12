// File: Date.cpp
// Date implementation file
//

#include <ctype.h>
#include <stdio.h>
#include "splib.h"
#include "splibint.h"

namespace splib {

Date::Date() {
    year = 0;
    month = 1;
    day = 1;
}

Date::Date(int y, int m, int d) {
    if (!validate(y, m, d)) {
        throw IllegalArgumentException();
    }
    year = y;
    month = m;
    day = d;
}

Date::Date(const _TCHAR* date) {
    int y;
    int m;
    int d;
    if (!parse(date, y, m, d)) {
        throw IllegalArgumentException();
    }
    year = y;
    month = m;
    day = d;
}

Date::~Date() {
}

bool Date::operator == (const Date& date) const {
    return year == date.year && month == date.month && day == date.day;
}

bool Date::operator != (const Date& date) const {
    return !(this->operator ==(date));
}

int Date::getYear() const {
    return year;
}

void Date::setYear(int y) {
    if (!validate(y, month, day)) {
        throw IllegalArgumentException();
    }
    year = y;
}

int Date::getMonth() const {
    return month;
}

void Date::setMonth(int m) {
    if (!validate(year, m, day)) {
        throw IllegalArgumentException();
    }
    month = m;
}

int Date::getDay() const {
    return day;
}

void Date::setDay(int d) {
    if (!validate(year, month, d)) {
        throw IllegalArgumentException();
    }
    day = d;
}

const _TCHAR* Date::getString() const {
    _TCHAR s[11];
#pragma warning (disable : 4996)
    _stprintf(s, _T("%04d-%02d-%02d"), year, month, day);
#pragma warning (default : 4996)
    Date* p = (Date*)this;
    p->str = s;
    return str.c_str();
}

void Date::setString(const _TCHAR* date) {
    int y;
    int m;
    int d;
    if (!parse(date, y, m, d)) {
        throw IllegalArgumentException();
    }
    year = y;
    month = m;
    day = d;
}

bool Date::validate(int year, int month, int day) {
    if (year < 0 || year > 9999) {
        return false;
    }
    if (month < 1 || month > 12) {
        return false;
    }
    int upperBound = 31;
    if (month == 4 || month == 6 || month == 9 || month == 11) {
        upperBound = 30;
    }
    if (month == 2) {
        // check for leap years
        if ((year % 400 == 0) || ((year % 4 == 0) && (year % 100 != 0))) {
            upperBound = 29;
        } else {
            upperBound = 28;
        }
    }
    if (day < 1 || day > upperBound) {
        return false;
    }
    return true;
}

bool Date::parse(const _TCHAR* date, int& year, int& month, int& day) {
    if (date == 0) {
        return false;
    }
    if (!_istdigit(date[0])
            || !_istdigit(date[1])
            || !_istdigit(date[2])
            || !_istdigit(date[3])
            || date[4] != '-'
            || !_istdigit(date[5])
            || !_istdigit(date[6])
            || date[7] != '-'
            || !_istdigit(date[8])
            || !_istdigit(date[9])) {
        return false;
    }
    int y;
    int m;
    int d;
#pragma warning (disable : 4996)
    _stscanf(date, _T("%d-%d-%d"), &y, &m, &d);
#pragma warning (default : 4996)
    if (!validate(y, m, d)) {
        return false;
    }
    year = y;
    month = m;
    day = d;
    return true;
}

}
