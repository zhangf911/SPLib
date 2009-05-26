// File: IndexLimits.h
// IndexLimits declaration file
//

#ifndef INDEXLIMITS_H
#define INDEXLIMITS_H

#include "splib.h"

namespace splib {

/**
 * A single point of validation of row and column indices
 * against their limits.
 */
class IndexLimits {
    public:
        /**
         * Validates a pair of indices against limits.
         * @param column column index
         * @param row row index
         * @return true if both indices are valid; false otherwise
         */
        static bool validate(int column, int row);

        /**
         * Validates a row index against the limits.
         * @param row row index
         * @return true if the row index is valid; false otherwise
         */
        static bool validateRow(int row);

        /**
         * Validates a column index against the limits.
         * @param column column index
         * @return true if the column index is valid; false otherwise
         */
        static bool validateColumn(int column);
};

}

#endif // INDEXLIMITS_H
