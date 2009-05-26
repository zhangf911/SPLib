// File: XlsxWriter.cpp
// XlsxWriter implementation file
//

#include "splib.h"
#include "XlsxWriterImpl.h"
#include "splibint.h"

namespace splib {

void XlsxWriter::write(Spreadsheet& sp, const _TCHAR* pathname) {
    XlsxWriterImpl::write(sp, pathname);
}

}
