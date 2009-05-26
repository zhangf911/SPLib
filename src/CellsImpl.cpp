// File: CellsImpl.cpp
// CellsImpl implementation file
//

#include "CellsImpl.h"
#include "IndexLimits.h"
#include "Util.h"
#include "splibint.h"

namespace splib {

Cell& CellsImpl::get(int index) {
    if (!IndexLimits::validateColumn(index)) {
        throw IllegalArgumentException();
    }
    return collection.get(index);
}

bool CellsImpl::contains(int index) {
    if (!IndexLimits::validateColumn(index)) {
        throw IllegalArgumentException();
    }
    return collection.contains(index);
}

void CellsImpl::remove(int index) {
    if (!IndexLimits::validateColumn(index)) {
        throw IllegalArgumentException();
    }
    return collection.remove(index);
}

int CellsImpl::size() const {
    return collection.size();
}

bool CellsImpl::isEmpty() const {
    return collection.isEmpty();
}

CellsImpl::Iterator* CellsImpl::iterator() {
    return collection.iterator();
}

CellsImpl::Entry CellsImpl::first() {
    return collection.first();
}

CellsImpl::Entry CellsImpl::last() {
    return collection.last();
}

}
