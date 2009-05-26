// File: XlsxWriterImpl.h
// XlsxWriterImpl declaration file
//

#ifndef XLSXWRITERIMPL_H
#define XLSXWRITERIMPL_H

#include "splib.h"

namespace splib {

/**
 * This class provides a static method that writes a spreadsheet
 * to a file in xlsx format.
 */
class XlsxWriterImpl {

    public:
        /** Writes a spreadsheet to a file in xlsx format. */
        static void write(Spreadsheet& sp, const _TCHAR* pathname);

    private:
        /** Writes the [Content_Types].xml entry of the xlsx package. */
        static void writeContentTypes(Spreadsheet& sp, class ZipArchive& ar);
        
        /** Writes the _rel/.rels entry of the xlsx package. */
        static void writeRels(ZipArchive& ar);
        
        /** Writes the docProps/app.xml entry of the xlsx package. */
        static void writeAppDocProps(Spreadsheet& sp, ZipArchive& ar);
        
        /** Writes the docProps/core.xml entry of the xlsx package. */
        static void writeCoreDocProps(ZipArchive& ar);
        
        /** Writes the xl/styles.xml entry of the xlsx package. */
        static void writeStyles(ZipArchive& ar);
        
        /** Writes XF entries to the current entry. */
        static void writeXFs(ZipArchive& ar);

        /**
         * Determines the index to an XF record for a given cell. This is
         * coupled with writeXFs().
         */
        static int xfIndex(Cell& cell);
        
        /** Writes the xl/_rels/workbook.xml.rels entry of the xlsx package. */
        static void writeWorkbookRels(Spreadsheet& sp, ZipArchive& ar);
        
        /** Writes the xl/workbook.xml entry of the xlsx package. */
        static void writeWorkbook(Spreadsheet& sp, ZipArchive& ar);
        
        /** Writes the xl/theme/theme1.xml entry of the xlsx package. */
        static void writeTheme(ZipArchive& ar);
        
        /** Writes a sheet entry to an xlsx package. */
        static void writeSheet(Table& table, int id, ZipArchive& ar);
        
        /** Writes a cell to the current zip entry. */
        static void writeCell(Cell& cell, int col, int row, ZipArchive& ar);
};

}

#endif // XLSXWRITERIMPL_H
