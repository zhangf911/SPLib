// File: SpreadsheetImpl.cpp
// SpreadsheetImpl implementation file
//

#include "splib.h"
#include "splibint.h"

namespace splib {

ExceptionImpl::ExceptionImpl() {
}

ExceptionImpl::ExceptionImpl(const _TCHAR* message) {
    msg = message;
}

const _TCHAR* ExceptionImpl::message() {
    return msg.c_str();
}

}
