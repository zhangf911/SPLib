// File: SpreadsheetImpl.cpp
// SpreadsheetImpl implementation file
//

#include "splib.h"
#include "TableImpl.h"
#include "splibint.h"

#include <string.h>

namespace splib {

SpreadsheetImpl::SpreadsheetImpl() {
}

SpreadsheetImpl::~SpreadsheetImpl() {
    for (std::vector<Table*>::size_type i = 0; i < tables.size(); i++) {
        delete tables[i];
    }
}

Table& SpreadsheetImpl::insertTable(int index, const _TCHAR* name) {
    if (index < 0 || index > tableCount()) {
        throw IllegalArgumentException();
    }
    if (name == 0) {
        throw IllegalArgumentException();
    }
    if (tableExists(name)) {
        throw IllegalArgumentException();
    }
    Table* table = new TableImpl(this);
    table->setName(name);
    tables.insert(tables.begin() + index, table);
    return *table;
}

Table& SpreadsheetImpl::table(int index) {
    if (index < 0 || index >= tableCount()) {
        throw IllegalArgumentException();
    }
    return *(tables[index]);
}

Table& SpreadsheetImpl::table(const _TCHAR* name) {
    if (name == 0) {
        throw IllegalArgumentException();
    }
    int index = find(name);
    if (index == -1) {
        throw IllegalArgumentException();
    }
    return *(tables[index]);
}

int SpreadsheetImpl::tableCount() const {
    return (int) tables.size();
}

bool SpreadsheetImpl::tableExists(const _TCHAR* name) {
    if (name == 0) {
        throw IllegalArgumentException();
    }
    return find(name) != -1;
}

void SpreadsheetImpl::removeTable(int index) {
    if (index < 0 || index >= tableCount()) {
        throw IllegalArgumentException();
    }
    delete tables[index];
    tables.erase(tables.begin() + index);
}

void SpreadsheetImpl::removeTable(const _TCHAR* name) {
    if (name == 0) {
        throw IllegalArgumentException();
    }
    int index = find(name);
    if (index == -1) {
        throw IllegalArgumentException();
    }
    removeTable(index);
}

int SpreadsheetImpl::find(const _TCHAR* name) {
    for (int i = 0; i < (int) tables.size(); i++) {
        if (_tcscmp(name, tables[i]->getName()) == 0) {
            return i;
        }
    }
    return -1;
}

}
