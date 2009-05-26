// File: IndexLimits.cpp
// IndexLimits implementation file
//

#include "IndexLimits.h"
#include "splibint.h"

namespace splib {

bool IndexLimits::validate(int column, int row) {
    if (!validateRow(row)) {
        return false;
    }
    if (!validateColumn(column)) {
        return false;
    }
    return true;
}

bool IndexLimits::validateRow(int row) {
    return row >= 0 && row <= 65535;
}

bool IndexLimits::validateColumn(int column) {
    return column >= 0 && column <= 255;
}

}
