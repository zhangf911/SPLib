// File: ColumnImpl.cpp
// ColumnImpl implementation file
//

#include "ColumnImpl.h"
#include "splibint.h"

namespace splib {

ColumnImpl::ColumnImpl() {
    width = -1;
}

double ColumnImpl::getWidth() const {
    return width;
}

void ColumnImpl::setWidth(double w) {
    width = w;
}

}
