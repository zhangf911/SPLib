// File: Strings.h
// Strings declaration file
//

#ifndef STRINGS_H
#define STRINGS_H

#include "splib.h"

namespace splib {

/** _TCHAR string. */
typedef std::basic_string<_TCHAR> _tstring;

/** Utility methods for manipulating strings. */
class Strings {
    public:
        /** Prepares a string for output to XML. */
        static _tstring xmlize(const _TCHAR* str);

        /** Replaces all occurrences of a substring within a string. */
        static _tstring replace(const _TCHAR* str,
                                const _TCHAR* substring,
                                const _TCHAR* replacement);
};

}

#endif // STRINGS_H
