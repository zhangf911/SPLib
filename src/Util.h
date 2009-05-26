// File: Util.h
// Util declaration file
//

#ifndef UTIL_H
#define UTIL_H

#include "splib.h"

namespace splib {

/** Static methods for parsing and producing cell locations. */
class Util {
    public:
        /**
         * Parses a cell location string, e.g. "A10".
         * @param s a pointer to the string to parse
         * @param col on successful exit, parsed column index of the cell;
         *        on failure, not modified
         * @param row on successful exit, parsed row index of the cell;
         *        on failure, not modified
         * @return true if the cell location is parsed successfully,
         *         false otherwise
         */
        static bool parseLocation(const _TCHAR* s, int& col, int& row);

        /**
         * Parses a cell location substring embedded in a bigger string.
         * @param s on enter, a pointer to the fragment to parse;
         *        on successful exit, a pointer to the position after
         *        the parsed fragment; on failure, not modified
         * @param col on successful exit, parsed column index of the cell;
         *        on failure, not modified
         * @param row on successful exit, parsed row index of the cell;
         *        on failure, not modified
         * @return true if the cell location is parsed successfully,
         *         false otherwise
         */
        static bool parseEmbeddedLocation(const _TCHAR*& s, int& col, int& row);

        /**
         * Builds a location string from column and row indices.
         * @param col column index
         * @param row row index
         * @return cell location string, e.g. "A10".
         */
        static std::basic_string<char> buildLocation(int col, int row);

        /**
         * Parses a column specification string, for example, "AA".
         * @param s a pointer to the string to parse
         * @param col on successful exit, parsed column index;
         *        on failure, not modified
         * @return true if the string is parsed successfully,
         *         false otherwise
         */
        static bool parseColumn(const _TCHAR* s, int& col);
};

}

#endif // UTIL_H
