// File: Formulas.h
// Formulas declaration file
//

#ifndef FORMULAS_H
#define FORMULAS_H

#include "splib.h"

namespace splib {

/** Static methods for formula parsing. */
class Formulas {
    public:
        /**
         * Parses a formula string. Only sums of cell ranges are
         * supported, such as "SUM(A1:B10)".
         * @param formula string to parse
         * @param c1 on successful exit, column index of the beginning of
         *        the summation range; on failure, not modified
         * @param r1 on successful exit, row index of the beginning of the
         *        summation range; on failure, not modified
         * @param c2 on successful exit, column index of the end of the
         *        summation range; on failure, not modified
         * @param r2 on successful exit, row index of the end of the
         *        summation range; on failure, not modified
         * @return true if the formula is parsed successfully, false
         *         otherwise
         */
        static bool parse(const _TCHAR* formula, int& c1, int& r1,
                          int& c2, int& r2);

    private:
        /**
         * Verifies that a pointer points to a specific character sequence
         * and advances the pointer to the position after the end of that
         * sequence.
         */
        static bool advance(const _TCHAR*& p, const _TCHAR* sequence);
};

}

#endif // FORMULAS_H
