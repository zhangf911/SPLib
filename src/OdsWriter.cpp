// File: OdsWriter.cpp
// OdsWriter implementation file
//

#include "splib.h"
#include "OdsWriterImpl.h"
#include "splibint.h"

namespace splib {

void OdsWriter::write(Spreadsheet& spreadsheet, const _TCHAR* pathname) {
    OdsWriterImpl::write(spreadsheet, pathname);
}

}
