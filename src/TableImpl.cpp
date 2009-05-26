// File: TableImpl.cpp
// TableImpl implementation file
//

#include "TableImpl.h"
#include "CellImpl.h"
#include "Util.h"
#include "IndexLimits.h"
#include "splibint.h"

#include <string.h>

namespace splib {

TableImpl::TableImpl(Spreadsheet* sp) : spreadsheet(sp) {
}

TableImpl::~TableImpl() {
}

Cell& TableImpl::cell(int column, int row) {
    if (!IndexLimits::validate(column, row)) {
        throw IllegalArgumentException();
    }
    return rows().get(row).cells().get(column);
}

Cell& TableImpl::cell(const _TCHAR* columnrow) {
    if (columnrow == 0) {
        throw IllegalArgumentException();
    }
    int column;
    int row;
    if (!Util::parseLocation(columnrow, column, row)) {
        throw IllegalArgumentException();
    }
    return cell(column, row);
}

void TableImpl::clearCell(int column, int row) {
    if (!IndexLimits::validate(column, row)) {
        throw IllegalArgumentException();
    }
    if (rows().contains(row)) {
        Row& r = rows().get(row);
        r.cells().remove(column);
        if (r.cells().isEmpty() && r.getHeight() < 0) {
            rows().remove(row);
        }
    }
}

void TableImpl::clearCell(const _TCHAR* columnrow) {
    if (columnrow == 0) {
        throw IllegalArgumentException();
    }
    int column;
    int row;
    if (!Util::parseLocation(columnrow, column, row)) {
        throw IllegalArgumentException();
    }
    clearCell(column, row);
}

void TableImpl::clearRange(int column1, int row1,
                           int column2, int row2) {
    if (!IndexLimits::validate(column1, row1)) {
        throw IllegalArgumentException();
    }
    if (!IndexLimits::validate(column2, row2)) {
        throw IllegalArgumentException();
    }
    int left;
    int right;
    int top;
    int bottom;
    if (column1 < column2) {
        left = column1;
        right = column2;
    } else {
        left = column2;
        right = column1;
    }
    if (row1 < row2) {
        top = row1;
        bottom = row2;
    } else {
        top = row2;
        bottom = row1;
    }
    for (int i = left; i <= right; i++) {
        for (int j = top; j <= bottom; j++) {
            clearCell(i, j);
        }
    }
}

void TableImpl::clearRange(const _TCHAR* columnrow1,
                           const _TCHAR* columnrow2) {
    if (columnrow1 == 0 || columnrow2 == 0) {
        throw IllegalArgumentException();
    }
    int column1;
    int row1;
    int column2;
    int row2;
    if (!Util::parseLocation(columnrow1, column1, row1)) {
        throw IllegalArgumentException();
    }
    if (!Util::parseLocation(columnrow2, column2, row2)) {
        throw IllegalArgumentException();
    }
    clearRange(column1, row1, column2, row2);
}

bool TableImpl::isEmptyCell(int column, int row) const {
    if (!IndexLimits::validate(column, row)) {
        throw IllegalArgumentException();
    }
    TableImpl* pThis = (TableImpl*)this;
    if (!pThis->rows().contains(row)) {
        return true;
    }
    Row& r = pThis->rows().get(row);
    return !r.cells().contains(column);
}

bool TableImpl::isEmptyCell(const _TCHAR* columnrow) const {
    if (columnrow == 0) {
        throw IllegalArgumentException();
    }
    int column;
    int row;
    if (!Util::parseLocation(columnrow, column, row)) {
        throw IllegalArgumentException();
    }
    return isEmptyCell(column, row);
}

int TableImpl::firstColumn() const {
    int res = -1;
    TableImpl* pThis = (TableImpl*)this;
    Rows::Iterator* i = pThis->rows().iterator();
    while (i->hasNext()) {
        Rows::Entry entry = i->next();
        Cells& cells = entry.object().cells();
        if (cells.size() > 0) {
            int first = cells.first().index();
            if (res == -1 || first < res) {
                res = first;
            }
        }
    }
    delete i;
    return res;
}

int TableImpl::lastColumn() const {
    int res = -1;
    TableImpl* pThis = (TableImpl*)this;
    Rows::Iterator* i = pThis->rows().iterator();
    while (i->hasNext()) {
        Rows::Entry entry = i->next();
        Cells& cells = entry.object().cells();
        if (cells.size() > 0) {
            int last = cells.last().index();
            if (res == -1 || last > res) {
                res = last;
            }
        }
    }
    delete i;
    return res;
}

int TableImpl::firstRow() const {
    TableImpl* pThis = (TableImpl*)this;
    Rows& rows = pThis->rows();
    if (!rows.isEmpty()) {
        return rows.first().index();
    } else {
        return -1;
    }
}

int TableImpl::lastRow() const {
    TableImpl* pThis = (TableImpl*)this;
    Rows& rows = pThis->rows();
    if (!rows.isEmpty()) {
        return rows.last().index();
    } else {
        return -1;
    }
}

Rows& TableImpl::rows() {
    return rowsImpl;
}

Columns& TableImpl::columns() {
    return columnsImpl;
}

const _TCHAR* TableImpl::getName() const {
    return name.c_str();
}

void TableImpl::setName(const _TCHAR* name) {
    for (int i = 0; i < spreadsheet->tableCount(); i++) {
        Table& t = spreadsheet->table(i);
        if (&t != this && _tcscmp(t.getName(), name) == 0) {
            throw IllegalArgumentException();
        }
    }
    this->name = name;
}

}
