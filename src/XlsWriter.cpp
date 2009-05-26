// File: XlsWriter.cpp
// XlsWriter implementation file
//

#include "splib.h"
#include "XlsWriterImpl.h"
#include "splibint.h"

namespace splib {

void XlsWriter::write(Spreadsheet& spreadsheet, const _TCHAR* pathname) {
    XlsWriterImpl::write(spreadsheet, pathname);
}

}
