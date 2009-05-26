// File: CellImpl.cpp
// CellImpl implementation file
//

#include "CellImpl.h"
#include "Formulas.h"
#include "IndexLimits.h"
#include "splibint.h"

namespace splib {

CellImpl::CellImpl() {
    type = NONE;
    hAlignment = HADEFAULT;
    vAlignment = VADEFAULT;
}

CellImpl::~CellImpl() {
}

const _TCHAR* CellImpl::getText() const {
    if (type != TEXT) {
        throw IllegalStateException();
    }
    return text.c_str();
}

void CellImpl::setText(const _TCHAR* t) {
    type = TEXT;
    text = t ? t : _T("");
}

long CellImpl::getLong() const  {
    if (type != LONG) {
        throw IllegalStateException();
    }
    return value.longValue;
}

void CellImpl::setLong(long l)  {
    type = LONG;
    value.longValue = l;
}

double CellImpl::getDouble() const {
    if (type != DOUBLE) {
        throw IllegalStateException();
    }
    return value.doubleValue;
}

void CellImpl::setDouble(double d) {
    type = DOUBLE;
    value.doubleValue = d;
}

const Date& CellImpl::getDate() const {
    if (type != DATE) {
        throw IllegalStateException();
    }
    return date;
}

void CellImpl::setDate(const Date& d) {
    type = DATE;
    date = d;
}

const Time& CellImpl::getTime() const {
    if (type != TIME) {
        throw IllegalStateException();
    }
    return time;
}

void CellImpl::setTime(const Time& t) {
    type = TIME;
    time = t;
}

const _TCHAR* CellImpl::getFormula() const {
    if (type != FORMULA) {
        throw IllegalStateException();
    }
    return formula.c_str();
}

void CellImpl::setFormula(const _TCHAR* f) {
    int c1, r1, c2, r2;
    if (!Formulas::parse(f, c1, r1, c2, r2)) {
        throw IllegalArgumentException();
    }
    if (!IndexLimits::validate(c1, r1)) {
        throw IllegalArgumentException();
    }
    if (!IndexLimits::validate(c2, r2)) {
        throw IllegalArgumentException();
    }
    type = FORMULA;
    formula = f;
}

void CellImpl::clear() {
    type = NONE;
}

Cell::Type CellImpl::getType() const {
    return type;
}

CellImpl::HAlignment CellImpl::getHAlignment() const {
    return hAlignment;
}

void CellImpl::setHAlignment(HAlignment h) {
    hAlignment = h;
}

CellImpl::VAlignment CellImpl::getVAlignment() const {
    return vAlignment;
}

void CellImpl::setVAlignment(VAlignment v) {
    vAlignment = v;
}

}
