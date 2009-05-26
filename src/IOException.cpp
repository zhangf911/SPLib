// File: IOException.cpp
// IOException implementation file
//

#include "splib.h"
#include "splibint.h"

namespace splib {

IOException::IOException() {
}

IOException::IOException(const _TCHAR* message) : ExceptionImpl(message) {
}

}
