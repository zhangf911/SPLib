// File: RowsImpl.cpp
// RowsImpl implementation file
//

#include "RowsImpl.h"
#include "IndexLimits.h"
#include "splibint.h"

namespace splib {

Row& RowsImpl::get(int index) {
    if (!IndexLimits::validateRow(index)) {
        throw IllegalArgumentException();
    }
    return collection.get(index);
}

bool RowsImpl::contains(int index) {
    if (!IndexLimits::validateRow(index)) {
        throw IllegalArgumentException();
    }
    return collection.contains(index);
}

void RowsImpl::remove(int index) {
    if (!IndexLimits::validateRow(index)) {
        throw IllegalArgumentException();
    }
    return collection.remove(index);
}

int RowsImpl::size() const {
    return collection.size();
}

bool RowsImpl::isEmpty() const {
    return collection.isEmpty();
}

RowsImpl::Iterator* RowsImpl::iterator() {
    return collection.iterator();
}

RowsImpl::Entry RowsImpl::first() {
    return collection.first();
}

RowsImpl::Entry RowsImpl::last() {
    return collection.last();
}

}
