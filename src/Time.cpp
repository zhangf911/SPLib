// File: Time.cpp
// Time implementation file
//

#include <ctype.h>
#include <stdio.h>
#include "splib.h"
#include "splibint.h"

namespace splib {

Time::Time() {
    hours = 0;
    minutes = 0;
    seconds = 0;
    millis = 0;
}

Time::Time(int h, int m, int s /*= 0*/, int ms /*= 0*/) {
    if (!validate(h, m, s, ms)) {
        throw IllegalArgumentException();
    }
    hours = h;
    minutes = m;
    seconds = s;
    millis = ms;
}

Time::Time(const _TCHAR* time) {
    int h;
    int m;
    int s;
    int ms;
    if (!parse(time, h, m, s, ms)) {
        throw IllegalArgumentException();
    }
    hours = h;
    minutes = m;
    seconds = s;
    millis = ms;
}

Time::~Time() {
}

bool Time::operator == (const Time& time) const {
    return hours == time.hours && minutes == time.minutes
        && seconds == time.seconds && millis == time.millis;
}

bool Time::operator != (const Time& time) const {
    return !(this->operator ==(time));
}

int Time::getHours() const {
    return hours;
}

void Time::setHours(int h) {
    if (h < 0 || h > 23) {
        throw IllegalArgumentException();
    }
    hours = h;
}

int Time::getMinutes() const {
    return minutes;
}

void Time::setMinutes(int m) {
    if (m < 0 || m > 59) {
        throw IllegalArgumentException();
    }
    minutes = m;
}

int Time::getSeconds() const {
    return seconds;
}

void Time::setSeconds(int s) {
    if (s < 0 || s > 59) {
        throw IllegalArgumentException();
    }
    seconds = s;
}

int Time::getMillis() const {
    return millis;
}

void Time::setMillis(int ms) {
    if (ms < 0 || ms > 999) {
        throw IllegalArgumentException();
    }
    millis = ms;
}

const _TCHAR* Time::getString() const {
    _TCHAR s[13];
#pragma warning (disable : 4996)
    _stprintf(s, _T("%02d:%02d:%02d.%03d"), hours, minutes, seconds, millis);
#pragma warning (default : 4996)
    Time* p = (Time*)this;
    p->str = s;
    return str.c_str();
}

void Time::setString(const _TCHAR* time) {
    int h;
    int m;
    int s;
    int ms;
    if (!parse(time, h, m, s, ms)) {
        throw IllegalArgumentException();
    }
    hours = h;
    minutes = m;
    seconds = s;
    millis = ms;
}

bool Time::validate(int hours, int minutes, int seconds, int millis) {
    if (hours < 0 || hours > 23) {
        return false;
    }
    if (minutes < 0 || minutes > 59) {
        return false;
    }
    if (seconds < 0 || seconds > 59) {
        return false;
    }
    if (millis < 0 || millis > 999) {
        return false;
    }
    return true;
}

bool Time::parse(const _TCHAR* time, int& hours, int& minutes, int& seconds,
                 int& millis) {
    if (time == 0) {
        return false;
    }
    if (!_istdigit(time[0])
            || !_istdigit(time[1])
            || time[2] != ':'
            || !_istdigit(time[3])
            || !_istdigit(time[4])) {
        return false;
    }
    int h;
    int m;
    int s = 0;
    int ms = 0;
#pragma warning (disable : 4996)
    _stscanf(time, _T("%d:%d"), &h, &m);
#pragma warning (default : 4996)

    if (time[5] != 0 && time[5] != ':') {
        return false;
    }
    if (time[5] == ':') {
        if (!_istdigit(time[6])
                || !_istdigit(time[7])) {
            return false;
        }
#pragma warning (disable : 4996)
        _stscanf(time + 6, _T("%d"), &s);
#pragma warning (default : 4996)

        if (time[8] != 0 && time[8] != '.') {
            return false;
        }
        if (time[8] == '.') {
            for (int i = 0; i < 3; i++) {
                if (!_istdigit(time[9 + i])) {
                    return false;
                }
            }
            if (time[12] != 0) {
                return false;
            }
#pragma warning (disable : 4996)
            _stscanf(time + 9, _T("%d"), &ms);
#pragma warning (default : 4996)
        }
    }

    if (!validate(h, m, s, ms)) {
        return false;
    }
    hours = h;
    minutes = m;
    seconds = s;
    millis = ms;
    return true;
}

}
