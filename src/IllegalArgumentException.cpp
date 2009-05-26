// File: IllegalArgumentException.cpp
// IllegalArgumentException implementation file
//

#include "splib.h"
#include "splibint.h"

namespace splib {

IllegalArgumentException::IllegalArgumentException() {
}

IllegalArgumentException::IllegalArgumentException(const _TCHAR* argument) 
    : ExceptionImpl(argument) {
}

}
