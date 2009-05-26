// File: OdsWriterImpl.h
// OdsWriterImpl declaration file
//

#ifndef ODSWRITERIMPL_H
#define ODSWRITERIMPL_H

#include "splib.h"

namespace splib {

/**
 * This class provides a static method that writes a spreadsheet
 * to a file in ods format.
 */
class OdsWriterImpl {

    public:
        /** Writes a spreadsheet to a file in ods format. */
        static void write(Spreadsheet& spreadsheet, const _TCHAR* pathname);

    private:
        /** Writes the META-INF/manifest.xml entry of the ods package. */
        static void writeManifest(class ZipArchive& ar);

        /** Writes the content.xml entry of the ods package. */
        static void writeContent(Spreadsheet& spreadsheet, ZipArchive& ar);

        /** Writes the meta.xml entry of the ods package. */
        static void writeMeta(ZipArchive& ar);

        /** Writes the mimetype entry of the ods package. */
        static void writeMimetype(ZipArchive& ar);

        /** Writes the settings.xml entry of the ods package. */
        static void writeSettings(ZipArchive& ar);

        /** Writes the styles.xml entry of the ods package. */
        static void writeStyles(ZipArchive& ar);

        /** Writes a table into the current zip entry. */
        static void writeTable(Table& table, int& nextColumnStyle,
            int& nextRowStyle, ZipArchive& ar);

        /** Writes all required row styles into the current zip entry. */
        static void writeRowStyles(Spreadsheet& sp, int& nextRowStyle,
            ZipArchive& ar);

        /** Writes all required column styles into the current zip entry. */
        static void writeColumnStyles(Spreadsheet& sp, int& nextColumnStyle,
            ZipArchive& ar);

        /** Writes table columns into the current zip entry. */
        static void writeColumns(Table& table, int& nextColumnStyle,
            ZipArchive& ar);

        /**
         * Writes the specified number of empty columns into the current
         * zip entry.
         */
        static void writeEmptyColumns(int columns, ZipArchive& ar);

        /**
         * Writes the specified number of empty rows into the current
         * zip entry.
         */
        static void writeEmptyRows(int rows, ZipArchive& ar);

        /** Writes a non-empty cell into the current zip entry. */
        static void writeCell(Cell& cell, ZipArchive& ar);

        /**
         * Writes the specified number of empty cells into the current
         * zip entry.
         */
        static void writeEmptyCells(int cells, ZipArchive& ar);

        /** Converts a Date object to a string in OpenDocument format. */
        static std::basic_string<char> date(const Date& date);

        /** Converts a Time object to a string in OpenDocument format. */
        static std::basic_string<char> time(const Time& time);

        /** Converts a cell formula to a string in OpenDocument format. */
        static std::basic_string<char> formula(const _TCHAR* formula);

        /** Writes cell styles into the current zip entry. */
        static void writeCellStyles(ZipArchive& ar);

        /** Determines the style index for a cell. */
        static int cellStyleIndex(Cell& cell);
};

}

#endif // ODSWRITERIMPL_H
