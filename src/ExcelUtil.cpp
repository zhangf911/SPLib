// File: ExcelUtil.cpp
// ExcelUtil implementation file
//

#include "ExcelUtil.h"
#include "splibint.h"

namespace splib {

double ExcelUtil::columnWidthUnits(double width) {
    return width * 42.85546875 / 225;
}

long ExcelUtil::date(const Date& date) {
    int nYear = date.getYear();
    int nMonth = date.getMonth();
    int nDay = date.getDay();

    // Excel/Lotus 123 have a bug with 29-02-1900. 1900 is not a
    // leap year, but Excel/Lotus 123 think it is...
    if (nDay == 29 && nMonth == 02 && nYear==1900) {
        return 60;
    }

    // DMY to Modified Julian calculatie with an extra substraction of 2415019.
    long nSerialDate = 
        int((1461 * (nYear + 4800 + int((nMonth - 14) / 12))) / 4) +
        int((367 * (nMonth - 2 - 12 * ((nMonth - 14) / 12))) / 12) -
        int((3 * (int((nYear + 4900 + int((nMonth - 14) / 12)) / 100))) / 4)
        + nDay - 2415019 - 32075;

    if (nSerialDate < 60) {
        // Because of the 29-02-1900 bug, any serial date 
        // under 60 is one off... Compensate.
        nSerialDate--;
    }

    return nSerialDate;
}

double ExcelUtil::time(const Time& time) {
    int h = time.getHours();
    int m = time.getMinutes();
    int s = time.getSeconds();
    int ms = time.getMillis();
    const static double msInDay = 24 * 60 * 60 * 1000;
    return (ms + s * 1000 + m * 60 * 1000 + h * 60 * 60 * 1000) / msInDay;
}

}
