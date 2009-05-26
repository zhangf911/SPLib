// File: XlsWriterImpl.cpp
// XlsWriterImpl implementation file
//

#include "splib.h"
#include "XlsWriterImpl.h"
#include "ByteArray.h"
#include "BasicExcel.h"
#include "ToUTF16.h"
#include "ExcelUtil.h"
#include "Formulas.h"
#include "splibint.h"

namespace splib {

void XlsWriterImpl::write(Spreadsheet& spreadsheet, const _TCHAR* pathname) {
    ByteArray byteArray;
    write(spreadsheet, byteArray);
    outputCompoundFile(byteArray, pathname);
}

void XlsWriterImpl::write(Spreadsheet& sp, ByteArray& out) {
    // BOF 0x0809
    byte BOF[] = {
        0x00, 0x06, 0x05, 0x00, 0xF2, 0x15, 0xCC, 0x07,
        0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00};
    writeRecord(0x0809, sizeof(BOF), BOF, out);    
    // WINDOW1 0x003D
    byte WINDOW1[] = {
        0x68, 0x01, 0x0E, 0x01, 0x5C, 0x3A, 0xBE, 0x23,
        0x38, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00,
        0x58, 0x02};
    writeRecord(0x003D, sizeof(WINDOW1), WINDOW1, out);
    // write required XF records
    writeXFs(out);
    // write a BOUNDSHEET record for each table
    std::vector<int> sheetRefOffsets;
    for (int i = 0; i < sp.tableCount(); i++) {
        Table& table = sp.table(i);
        // record offset to reference to sheet streams for future filling
        int sheetRefOffset = writeBoundsheet(table.getName(), out);
        sheetRefOffsets.push_back(sheetRefOffset);
    }
    // EOF
    writeRecord(0x000A, 0, 0, out);
    // write each table as a sheet stream
    for (int i = 0; i < sp.tableCount(); i++) {
        // fill reference to this sheet first
        write4bytes(out.size(), out.data() + sheetRefOffsets[i]);
        Table& table = sp.table(i);
        writeTable(table, out);
    }
}

void XlsWriterImpl::outputCompoundFile(const ByteArray& contents,
                                       const _TCHAR* pathname) {
    const char* data = (const char*)contents.data();
    int size = contents.size();
    YCompoundFiles::CompoundFile file;
    if (!file.Create(pathname)) {
        throw IOException(_T("error creating compound file"));
    }
    file.MakeFile(_T("Workbook"));
    if (file.WriteFile(_T("Workbook"), (const char*)data, size)
            != YCompoundFiles::CompoundFile::SUCCESS) {
        throw IOException(_T("error writing workbook stream"));
    }
    if (!file.Close()) {
        throw IOException(_T("error closing compound file"));
    }
}

void XlsWriterImpl::writeXFs(ByteArray& out) {
    // XF 0x00E0
    // Offset   Size    Contents
    // 0        2       Index to FONT record
    // 2        2       Index to FORMAT record
    // 4        2       Bit     Mask    Contents
    //                  2-0     0007H   XF_TYPE_PROT – XF type, cell prot
    //                  15-4    FFF0H   Index to parent style XF
    //                                  (always FFFH in style XFs)
    // 6        1       Bit     Mask    Contents
    //                  2-0     07H     XF_HOR_ALIGN – Horizontal alignment
    //                  3       08H     1 = Text is wrapped at right border
    //                  6-4     70H     XF_VERT_ALIGN – Vertical alignment
    // 7        1       XF_ROTATION: Text rotation angle
    // 8        1       Indentation and direction
    // 9        1       Bit     Mask    Contents
    //                  7-2     FCH     XF_USED_ATTRIB – Used attributes
    // 10       4       Cell border lines and background area
    // 14       4       Various graphics options
    // 18       2       Colors
    byte XF[] = {
        0x00, 0x00, 0x00, 0x00, 0xF5, 0xFF, 0x20, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00};
    // write 16 basic records
    for (int i = 0; i < 16; i++) {
        writeRecord(0x00E0, sizeof(XF), XF, out);
    }
    // write XF records that are going to be used
    const int FORMAT_OFFSET = 2;
    const int ALIGN_OFFSET = 6;
    const ushort FORMAT_GENERAL = 0;
    const ushort FORMAT_DATE = 14;
    const ushort FORMAT_TIME = 21;
    const byte HALIGN_DEFAULT = 0x00;
    const byte HALIGN_LEFT = 0x01;
    const byte HALIGN_CENTERED = 0x02;
    const byte HALIGN_RIGHT = 0x03;
    const byte HALIGN_JUSTIFIED = 0x05;
    const byte HALIGN_FILLED = 0x04;
    const byte VALIGN_TOP = 0x00;
    const byte VALIGN_MIDDLE = 0x10;
    const byte VALIGN_BOTTOM = 0x20;
    ushort formats[] = {FORMAT_GENERAL, FORMAT_DATE, FORMAT_TIME};
    byte hAligns[] = {HALIGN_DEFAULT, HALIGN_LEFT, HALIGN_CENTERED,
        HALIGN_RIGHT, HALIGN_JUSTIFIED, HALIGN_FILLED};
    byte vAligns[] = {VALIGN_BOTTOM, VALIGN_TOP, VALIGN_MIDDLE};
    write2bytes(0, XF + 4); // XF_TYPE_PROT = 0, index to parent style XF = 0
    for (int i = 0; i < sizeof(formats) / sizeof(ushort); i++) {
        for (int j = 0; j < sizeof(hAligns); j++) {
            for (int k = 0; k < sizeof(vAligns); k++) {
                write2bytes(formats[i], XF + FORMAT_OFFSET);
                XF[ALIGN_OFFSET] = hAligns[j] | vAligns[k];
                writeRecord(0x00E0, sizeof(XF), XF, out);
            }
        }
    }
}

XlsWriterImpl::ushort XlsWriterImpl::xfIndex(Cell& cell) {
    int i = 0;
    switch (cell.getType()) {
        case Cell::DATE:        i = 1; break;
        case Cell::TIME:        i = 2; break;
    }
    int j = 0;
    switch (cell.getHAlignment()) {
        case Cell::LEFT:        j = 1; break;
        case Cell::CENTER:      j = 2; break;
        case Cell::RIGHT:       j = 3; break;
        case Cell::JUSTIFIED:   j = 4; break;
        case Cell::FILLED:      j = 5; break;
    }
    int k = 0;
    switch (cell.getVAlignment()) {
        case Cell::TOP:         k = 1; break;
        case Cell::MIDDLE:      k = 2; break;
    }
    return (ushort)(i * 3 * 6 + j * 3 + k + 16);
}

int XlsWriterImpl::writeBoundsheet(const _TCHAR* sheetName, ByteArray& out) {
    // BOUNDSHEET 0x0085
    // Offset   Size    Contents
    // 0        4       Sheet's BOF stream position, 0 is OK
    // 4        1       Visibility: 0 = visible
    // 5        1       Sheet type: 0 = worksheet
    // 6        var     Sheet name: UTF16 string, 8-bit string length
    ToUTF16 utf16(sheetName);
    ushort len = (ushort)(4 + 1 + 1 + 2 + utf16.length() * 2);
    writeRecordHeader(0x0085, len, out);
    int retval = out.size();
    write4bytes(0, out);
    out.write(0);
    out.write(0);
    out.write((byte)utf16.length());    // string length (8 bits)
    out.write(1);                       // option flags
    out.write(utf16.get(), utf16.length() * 2);
    return retval;
}

void XlsWriterImpl::writeTable(Table& table, ByteArray& out) {
    // BOF 0x0809
    byte BOF[] = {
        0x00, 0x06, 0x10, 0x00, 0xF2, 0x15, 0xCC, 0x07,
        0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x00};
    writeRecord(0x0809, sizeof(BOF), BOF, out);
    // columns
    writeColumns(table, out);
    // rows
    writeRows(table, out);    
    // EOF
    writeRecord(0x000A, 0, 0, out);
}

void XlsWriterImpl::writeColumns(Table& table, ByteArray& out) {
    Columns::Iterator* colIt = table.columns().iterator();
    while (colIt->hasNext()) {
        Columns::Entry entry = colIt->next();
        int col = entry.index();
        double width = entry.object().getWidth();
        if (width < 0) {
            continue;
        }
        // COLINFO 0x007D
        // Offset   Size    Contents
        // 0        2       Index to first column in the range
        // 2        2       Index to last column in the range
        // 4        2       Width of the columns in 1/256 of the width of
        //                  the zero character, using default font (first
        //                  FONT record in the file)
        // 6        2       Index to XF record for default column formatting
        // 8        2       Option flags:
        //                  Bits    Mask    Contents
        //                  0       0001H   1 = Columns are hidden
        //                  10-8    0700H   Outline level of the columns
        //                                  (0 = no outline)
        //                  12      1000H   1 = Columns are collapsed
        // 10       2       Not used
        byte COLINFO[] = {
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0F, 0x00,
            0x06, 0x00, 0x00, 0x00};
        write2bytes((ushort)col, COLINFO);
        write2bytes((ushort)col, COLINFO + 2);
        ushort w = (ushort)(ExcelUtil::columnWidthUnits(width)* 256);
        write2bytes(w, COLINFO + 4);
        writeRecord(0x007D, sizeof(COLINFO), COLINFO, out);
    }
    delete colIt;
}

void XlsWriterImpl::writeRows(Table& table, ByteArray& out) {
    Rows::Iterator* rowIt = table.rows().iterator();
    while (rowIt->hasNext()) {
        Rows::Entry entry = rowIt->next();
        int row = entry.index();
        double height = entry.object().getHeight();
        Cells& cells = entry.object().cells();
        if (cells.size() == 0 && height < 0) {
            continue;
        }
        ushort firstCell = 0;
        ushort lastCell = (ushort)-1;
        if (cells.size() > 0) {
            firstCell = (ushort)cells.first().index();
            lastCell = (ushort)cells.last().index();
        }
        // ROW 0x0208
        // Offset   Size    Contents
        // 0        2       Index of this row
        // 2        2       Index to column of the first cell
        // 4        2       Index to column of the last cell + 1
        // 6        2       Bit     Mask    Contents
        //                  14-0    7FFFH   Height of the row, in twips
        //                                  (= 1/20 of a point)
        //                  15      8000H   0 = Row has custom height;
        //                                  1 = Row has default height
        // 8        2       Not used
        // 10       2       Not used
        // 12       4       Option flags and default row formatting:
        //                  0x00000100 works just fine. 0x04 should be
        //                  added to apply custom height
        byte ROW[] = {
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x80,
            0x06, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00};
        write2bytes((ushort)row, ROW);
        write2bytes(firstCell, ROW + 2);
        write2bytes(lastCell + 1, ROW + 4);
        if (height >= 0) {
            write2bytes(((ushort)(height * 20)) & 0x7FFF, ROW + 6);
            write4bytes(0x00000140, ROW + 12);
        }
        writeRecord(0x0208, sizeof(ROW), ROW, out);
        // now the cells should go
        Cells::Iterator* cellIt = cells.iterator();
        while (cellIt->hasNext()) {
            Cells::Entry entry = cellIt->next();
            int col = entry.index();
            Cell& cell = entry.object();
            if (cell.getType() == Cell::NONE) {
                continue;
            }
            writeCell(cell, (ushort)col, (ushort)row, out);
        }
        delete cellIt;
    }
    delete rowIt;
}

void XlsWriterImpl::writeCell(Cell& cell, ushort col, ushort row,
                              ByteArray& out) {
    if (cell.getType() == Cell::TEXT) {
        // LABEL 0x0204
        // Offset   Size    Contents
        // 0        2       Index to row
        // 2        2       Index to column
        // 4        2       Index to XF record
        // 6        var.    Unicode string, 16-bit string length
        ToUTF16 utf16(cell.getText());
        writeRecordHeader(0x0204, (ushort)(6 + 3 + utf16.length() * 2), out);
        write2bytes(row, out);
        write2bytes(col, out);
        write2bytes(xfIndex(cell), out);
        write2bytes((ushort)utf16.length(), out);   // string length (16 bits)
        out.write(1);                               // option flags
        out.write(utf16.get(), utf16.length() * 2);
        return;
    }
    if (cell.getType() == Cell::LONG) {
        // NUMBER 0x0203 (long)
        // Offset   Size    Contents
        // 0        2       Index to row
        // 2        2       Index to column
        // 4        2       Index to XF record
        // 6        8       IEEE 754 floating-point value
        //                  (64-bit double precision)
        writeRecordHeader(0x0203, 14, out);
        write2bytes(row, out);
        write2bytes(col, out);
        write2bytes(xfIndex(cell), out);
        writeDouble(cell.getLong(), out);
        return;
    }
    if (cell.getType() == Cell::DOUBLE) {
        // NUMBER 0x0203 (double)
        // Offset   Size    Contents
        // 0        2       Index to row
        // 2        2       Index to column
        // 4        2       Index to XF record
        // 6        8       IEEE 754 floating-point value
        //                  (64-bit double precision)
        writeRecordHeader(0x0203, 14, out);
        write2bytes(row, out);
        write2bytes(col, out);
        write2bytes(xfIndex(cell), out);
        writeDouble(cell.getDouble(), out);
        return;
    }
    if (cell.getType() == Cell::DATE) {
        // NUMBER 0x0203 (date)
        // Offset   Size    Contents
        // 0        2       Index to row
        // 2        2       Index to column
        // 4        2       Index to XF record
        // 6        8       IEEE 754 floating-point value
        //                  (64-bit double precision)
        writeRecordHeader(0x0203, 14, out);
        write2bytes(row, out);
        write2bytes(col, out);
        write2bytes(xfIndex(cell), out);
        writeDouble(ExcelUtil::date(cell.getDate()), out);
        return;
    }
    if (cell.getType() == Cell::TIME) {
        // NUMBER 0x0203 (time)
        // Offset   Size    Contents
        // 0        2       Index to row
        // 2        2       Index to column
        // 4        2       Index to XF record
        // 6        8       IEEE 754 floating-point value
        //                  (64-bit double precision)
        writeRecordHeader(0x0203, 14, out);
        write2bytes(row, out);
        write2bytes(col, out);
        write2bytes(xfIndex(cell), out);
        writeDouble(ExcelUtil::time(cell.getTime()), out);
        return;
    }
    if (cell.getType() == Cell::FORMULA) {
        // FORMULA 0x0006
        // Offset   Size    Contents
        // 0        2       Index to row
        // 2        2       Index to column
        // 4        2       Index to XF record
        // 6        8       Result of the formula
        // 14       2       Option flags: 0x0002 = Calculate on open
        // 16       4       Not used
        // 20       var     Formula data (RPN token array)
        int c1, r1, c2, r2;
        Formulas::parse(cell.getFormula(), c1, r1, c2, r2);
        byte FORMULA[] = {
            0x0D, 0x00, 0x25, 0x01, 0x00, 0x02, 0x00, 0x03,
            0xC0, 0x04, 0xC0, 0x19, 0x10, 0x00, 0x00};
        write2bytes((ushort)r1, FORMULA + 3);
        write2bytes((ushort)r2, FORMULA + 5);
        FORMULA[7] = (byte)c1;
        FORMULA[9] = (byte)c2;
        ushort len = (ushort)(6 + 8 + 2 + 4 + sizeof(FORMULA));
        writeRecordHeader(0x0006, len, out);
        write2bytes(row, out);
        write2bytes(col, out);
        write2bytes(xfIndex(cell), out);
        writeDouble(0, out);
        write2bytes(0x02, out);
        write4bytes(0, out);
        out.write(FORMULA, sizeof(FORMULA));
        return;
    }
}

void XlsWriterImpl::writeRecord(ushort id, ushort len, const byte* data,
                                ByteArray& out) {
    writeRecordHeader(id, len, out);
    out.write(data, len);
}

void XlsWriterImpl::writeRecordHeader(ushort id, ushort len, ByteArray& out) {
    write2bytes(id, out);
    write2bytes(len, out);
}

void XlsWriterImpl::write2bytes(ushort value, ByteArray& out) {
    out.write((byte)(value & 0xFF));
    out.write((byte)((value >> 8) & 0xFF));
}

void XlsWriterImpl::write2bytes(ushort value, byte* out) {
    out[0] = (byte)(value & 0xFF);
    out[1] = (byte)((value >> 8) & 0xFF);
}

void XlsWriterImpl::write4bytes(ulong value, ByteArray& out) {
    out.write((byte)(value & 0xFF));
    out.write((byte)((value >> 8) & 0xFF));
    out.write((byte)((value >> 16) & 0xFF));
    out.write((byte)((value >> 24) & 0xFF));
}

void XlsWriterImpl::write4bytes(ulong value, byte* out) {
    out[0] = (byte)(value & 0xFF);
    out[1] = (byte)((value >> 8) & 0xFF);
    out[2] = (byte)((value >> 16) & 0xFF);
    out[3] = (byte)((value >> 24) & 0xFF);
}

void XlsWriterImpl::writeDouble(double value, ByteArray& out) {
    out.write((byte*)&value, 8);
}

}
