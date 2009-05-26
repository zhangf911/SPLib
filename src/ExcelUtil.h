// File: ExcelUtil.h
// ExcelUtil declaration file
//

#ifndef EXCELUTIL_H
#define EXCELUTIL_H

#include "splib.h"

namespace splib {

/** Excel-related utility methods. */
class ExcelUtil {
    public:
        /**
         * Ñonverts a width measured in points to a width measured in
         * characters of the average digit width of the default font.
         */
        static double columnWidthUnits(double width);

        /** Converts a Date object to a value in Excel format. */
        static long date(const Date& date);
        
        /** Converts a time object to a value in Excel format. */
        static double time(const Time& time);
};

}

#endif // EXCELUTIL_H
