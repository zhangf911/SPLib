// File: ColumnsImpl.cpp
// ColumnsImpl implementation file
//

#include "ColumnsImpl.h"
#include "IndexLimits.h"
#include "Util.h"
#include "splibint.h"

namespace splib {

Column& ColumnsImpl::get(int index) {
    if (!IndexLimits::validateColumn(index)) {
        throw IllegalArgumentException();
    }
    return collection.get(index);
}

Column& ColumnsImpl::get(const _TCHAR* column) {
    int col;
    if (!Util::parseColumn(column, col)) {
        throw IllegalArgumentException();
    }
    return get(col);
}

bool ColumnsImpl::contains(int index) {
    if (!IndexLimits::validateColumn(index)) {
        throw IllegalArgumentException();
    }
    return collection.contains(index);
}

void ColumnsImpl::remove(int index) {
    if (!IndexLimits::validateColumn(index)) {
        throw IllegalArgumentException();
    }
    return collection.remove(index);
}

int ColumnsImpl::size() const {
    return collection.size();
}

bool ColumnsImpl::isEmpty() const {
    return collection.isEmpty();
}

ColumnsImpl::Iterator* ColumnsImpl::iterator() {
    return collection.iterator();
}

ColumnsImpl::Entry ColumnsImpl::first() {
    return collection.first();
}

ColumnsImpl::Entry ColumnsImpl::last() {
    return collection.last();
}

}
