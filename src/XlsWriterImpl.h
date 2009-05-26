// File: XlsWriterImpl.h
// XlsWriterImpl declaration file
//

#ifndef XLSWRITERIMPL_H
#define XLSWRITERIMPL_H

#include "splib.h"

namespace splib {

/**
 * This class provides a static method that writes a spreadsheet
 * to a file in xls format.
 */
class XlsWriterImpl {

    public:
        /** Writes a spreadsheet to a file in xls format. */
        static void write(Spreadsheet& sp, const _TCHAR* pathname);

    private:
        /** The type represents a byte */
        typedef unsigned char byte;

        /** The type represents an unsigned short value */
        typedef unsigned short ushort;

        /** The type represents an unsigned long value */
        typedef unsigned long ulong;

        /**
         * Generates byte representation of a spreadsheet in Excel 97/2000
         * format.
         */
        static void write(Spreadsheet& sp, class ByteArray& out);

        /**
         * Outputs a byte array as an Excel workbook into an OLE compound
         * file.
         */
        static void outputCompoundFile(const ByteArray& contents,
            const _TCHAR* pathname);

        /** Generates necessary XF records. */
        static void writeXFs(ByteArray& out);

        /**
         * Determines the index to an XF record for a given cell. This is
         * coupled with writeXFs().
         */
        static ushort xfIndex(Cell& cell);

        /** Generates a BOUNDSHEET record for a sheet. */
        static int writeBoundsheet(const _TCHAR* sheetName, ByteArray& out);

        /** Generates byte representation of a table (worksheet) */
        static void writeTable(Table& table, ByteArray& out);

        /** Generates byte representation of table columns */
        static void writeColumns(Table& table, ByteArray& out);
        
        /** Generates byte representation of table rows */
        static void writeRows(Table& table, ByteArray& out);
        
        /** Generates byte representation of a cell */
        static void writeCell(Cell& cell, ushort col, ushort row,
            ByteArray& out);

        /** Writes a BIFF record to a byte array. */
        static void writeRecord(ushort id, ushort len, const byte* data,
            ByteArray& out);

        /** Writes a BIFF record header without the actual record */
        static void writeRecordHeader(ushort id, ushort len, ByteArray& out);

        /** Adds two bytes to a byte array */
        static void write2bytes(ushort value, ByteArray& out);

        /** Writes two bytes at a specified location */
        static void write2bytes(ushort value, byte* out);

        /** Adds four bytes to a byte array */
        static void write4bytes(ulong value, ByteArray& out);

        /** Writes four bytes at a specified location */
        static void write4bytes(ulong value, byte* out);

        /** Adds a double value to a byte array */
        static void writeDouble(double value, ByteArray& out);
};

}

#endif // XLSWRITERIMPL_H
