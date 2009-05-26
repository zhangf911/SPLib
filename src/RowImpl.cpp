// File: RowImpl.cpp
// RowImpl implementation file
//

#include "RowImpl.h"
#include "splibint.h"

namespace splib {

RowImpl::RowImpl() {
    height = -1;
}

double RowImpl::getHeight() const {
    return height;
}

void RowImpl::setHeight(double h) {
    height = h;
}

Cells& RowImpl::cells() {
    return cellsImpl;
}

}
