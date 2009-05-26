// File: IllegalStateException.cpp
// IllegalStateException implementation file
//

#include "splib.h"
#include "splibint.h"

namespace splib {

IllegalStateException::IllegalStateException() {
}

IllegalStateException::IllegalStateException(const _TCHAR* message)
    : ExceptionImpl(message) {
}

}
