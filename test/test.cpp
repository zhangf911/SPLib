// test.cpp : defines the entry point for the console application.
//

#include "splib.h"
#include <assert.h>
#include <math.h>
#include <sstream>
#include <stdio.h>
#include <string.h>

// these lines define the verify() macro that is equivalent to assert() in
// debug configurations but also works in release configurations; this is
// used to run these tests in release builds
#ifdef NDEBUG
#define verify(exp) (void)((exp) || (_verify(#exp, __FILE__, __LINE__), 0))
void _verify(const char *expr, const char *filename, unsigned lineno) {
    fprintf(stderr, "Verification failed: %s, file %s, line %d\n",
        expr, filename, lineno);
    fflush(stderr);
    abort();
}
#else // !NDEBUG
#define verify(exp) assert(exp)
#endif // NDEBUG

// some defines to allow this test to compile and run on both
// Windows and Linux
#ifndef WIN32
#define _T(x) x
#define _tmain  main
#define _tcscmp strcmp
#endif // WIN32

/**
 * Tests the Spreadsheet interface implementation.
 */
void testSpreadsheet();

/**
 * Tests the Table interface implementation.
 */
void testTable();

/**
 * Tests the Cell interface implementation.
 */
void testCell();

/**
 * Tests the Date class.
 */
void testDate();

/**
 * Tests the Time class.
 */
void testTime();

/**
 * Tests the XlsWriter, XlsxWriter, and OdsWriter classes.
 */
void testWriters();

/**
 * Setups a test spreadsheet for writer testing.
 */
void setupTestSpreadsheet(splib::Spreadsheet& sc);

/**
 * Setups a test table demonstrating different cell types.
 */
void setupCellTypesTable(splib::Spreadsheet& sc);

/**
 * Setups a test table demonstrating formulas.
 */
void setupFormulasTable(splib::Spreadsheet& sc);

/**
 * Setups a test table demonstrating alignment options.
 */
void setupAlignmentTable(splib::Spreadsheet& sc);

/**
 * Setups a test table that spans many rows and columns.
 */
void setupRowsAndColumnsTable(splib::Spreadsheet& sc);

/**
 * Setups a test table demonstrating special characters in strings.
 */
void setupSpecialCharsTable(splib::Spreadsheet& sc);

/**
 * Setups a test table demonstrating row heights and columns widths.
 */
void setupHeightsAndWidthsTable(splib::Spreadsheet& sc);

/**
 * Setups an empty table.
 */
void setupEmptyTable(splib::Spreadsheet& sc);


/**
 * Program entry point.
 */
int _tmain(int argc, _TCHAR* argv[]) {
    argc, argv;
    testSpreadsheet();
    testTable();
    testCell();
    testDate();
    testTime();
    testWriters();
}

void testSpreadsheet() {
    splib::SpreadsheetImpl sc;

    // insertTable(), table(), tableCount(), and tableExists()
    verify(!sc.tableExists(_T("table 0")));
    verify(!sc.tableExists(_T("table 1")));
    verify(!sc.tableExists(_T("table 2")));
    verify(!sc.tableExists(_T("table 3")));
    splib::Table& table1 = sc.insertTable(0, _T("table 1"));
    splib::Table& table0 = sc.insertTable(0, _T("table 0"));
    splib::Table& table2 = sc.insertTable(2, _T("table 2"));
    splib::Table& table3 = sc.insertTable(3, _T("table 3"));
    verify(sc.tableExists(_T("table 0")));
    verify(sc.tableExists(_T("table 1")));
    verify(sc.tableExists(_T("table 2")));
    verify(sc.tableExists(_T("table 3")));
    verify(sc.tableCount() == 4);
    verify(&table0 == &sc.table(0));
    verify(&table1 == &sc.table(1));
    verify(&table2 == &sc.table(2));
    verify(&table3 == &sc.table(3));
    verify(&table0 == &sc.table(_T("table 0")));
    verify(&table1 == &sc.table(_T("table 1")));
    verify(&table2 == &sc.table(_T("table 2")));
    verify(&table3 == &sc.table(_T("table 3")));

    // removeTable()
    sc.removeTable(1);
    verify(sc.tableExists(_T("table 0")));
    verify(!sc.tableExists(_T("table 1")));
    verify(sc.tableExists(_T("table 2")));
    verify(sc.tableExists(_T("table 3")));
    verify(sc.tableCount() == 3);
    verify(&table0 == &sc.table(0));
    verify(&table2 == &sc.table(1));
    verify(&table3 == &sc.table(2));
    verify(&table0 == &sc.table(_T("table 0")));
    verify(&table2 == &sc.table(_T("table 2")));
    verify(&table3 == &sc.table(_T("table 3")));

    sc.removeTable(_T("table 3"));
    sc.removeTable(_T("table 0"));    
    verify(!sc.tableExists(_T("table 0")));
    verify(!sc.tableExists(_T("table 1")));
    verify(sc.tableExists(_T("table 2")));
    verify(!sc.tableExists(_T("table 3")));
    verify(sc.tableCount() == 1);
    verify(&table2 == &sc.table(0));
    verify(&table2 == &sc.table(_T("table 2")));

    // Error handling
    try {
        sc.insertTable(2, _T("does not matter"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.insertTable(1, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.insertTable(1, _T("table 2"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.table(1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.table((_TCHAR*)0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.table(_T("table 1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.tableExists((_TCHAR*)0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.removeTable(1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.removeTable(-1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.removeTable((_TCHAR*)0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        sc.removeTable(_T("table 1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
}

void testTable() {
    splib::SpreadsheetImpl sc;
    splib::Table& table = sc.insertTable(0, _T("table"));

    // cell() and isEmptyCell()
    verify(table.isEmptyCell(0, 0));
    verify(table.isEmptyCell(_T("A1")));
    splib::Cell& cell_0_0 = table.cell(0, 0);
    splib::Cell& cell_A_1 = table.cell(_T("A1"));
    verify(!table.isEmptyCell(0, 0));
    verify(!table.isEmptyCell(_T("A1")));
    verify(&cell_0_0 == &cell_A_1);
    verify(&table.cell(0, 0) == &cell_0_0);
    verify(&table.cell(_T("A1")) == &cell_A_1);

    // cell() and cell position parsing
    verify(&table.cell(_T("A1")) == &table.cell(0, 0));
    verify(&table.cell(_T("B2")) == &table.cell(1, 1));
    verify(&table.cell(_T("C3")) == &table.cell(2, 2));
    verify(&table.cell(_T("Z1")) == &table.cell(25, 0));
    verify(&table.cell(_T("AA1")) == &table.cell(26, 0));
    verify(&table.cell(_T("AB2")) == &table.cell(27, 1));
    verify(&table.cell(_T("AC3")) == &table.cell(28, 2));
    verify(&table.cell(_T("BA1")) == &table.cell(52, 0));
    verify(&table.cell(_T("IV1")) == &table.cell(255, 0));
    verify(&table.cell(_T("IV65536")) == &table.cell(255, 65535));
    table.clearCell(_T("IV65536"));

    // clearCell()
    table.cell(_T("A1"));
    verify(!table.isEmptyCell(_T("A1")));
    table.cell(_T("B2"));
    verify(!table.isEmptyCell(_T("B2")));
    table.cell(_T("C3"));
    verify(!table.isEmptyCell(_T("C3")));
    table.cell(_T("A4"));    
    verify(!table.isEmptyCell(_T("A4")));
    table.clearCell(_T("A1"));
    verify(table.isEmptyCell(_T("A1")));
    table.clearCell(_T("B2"));
    verify(table.isEmptyCell(_T("B2")));
    table.clearCell(2, 2);
    verify(table.isEmptyCell(_T("C3")));
    table.clearCell(_T("A4"));    
    verify(table.isEmptyCell(_T("A4")));

    // clearRange()
    table.cell(_T("A1"));
    table.cell(_T("B2"));
    table.cell(_T("C3"));
    table.clearRange(_T("A1"), _T("C3"));
    verify(table.isEmptyCell(_T("A1")));
    verify(table.isEmptyCell(_T("B2")));
    verify(table.isEmptyCell(_T("C3")));

    table.cell(100,  999);
    table.cell( 99, 1000);
    table.cell(100, 1000);
    table.cell(110, 1100);
    table.cell(110, 1101);
    table.cell(111, 1100);
    table.clearRange(100, 1000, 110, 1100);
    verify(!table.isEmptyCell(100,  999));
    verify(!table.isEmptyCell( 99, 1000));
    verify( table.isEmptyCell(100, 1000));
    verify( table.isEmptyCell(110, 1100));
    verify(!table.isEmptyCell(110, 1101));
    verify(!table.isEmptyCell(111, 1100));
    table.clearRange(111, 1101, 99, 999); // non-normalized
    verify(table.isEmptyCell(100,  999));
    verify(table.isEmptyCell( 99, 1000));
    verify(table.isEmptyCell(110, 1101));
    verify(table.isEmptyCell(111, 1100));

    // firstColumn() & firstRow() & lastColumn() & lastRow()
    table.clearRange(table.firstColumn(), table.firstRow(),
        table.lastColumn(), table.lastRow());
    verify(table.firstColumn() == -1);
    verify(table.firstRow() == -1);
    verify(table.lastColumn() == -1);
    verify(table.lastRow() == -1);
    table.cell(5, 10);
    verify(table.firstColumn() == 5);
    verify(table.firstRow() == 10);
    verify(table.lastColumn() == 5);
    verify(table.lastRow() == 10);
    table.cell(5, 15);
    verify(table.firstColumn() == 5);
    verify(table.firstRow() == 10);
    verify(table.lastColumn() == 5);
    verify(table.lastRow() == 15);
    table.cell(5, 5);
    verify(table.firstColumn() == 5);
    verify(table.firstRow() == 5);
    verify(table.lastColumn() == 5);
    verify(table.lastRow() == 15);
    table.cell(3, 5);
    verify(table.firstColumn() == 3);
    verify(table.firstRow() == 5);
    verify(table.lastColumn() == 5);
    verify(table.lastRow() == 15);
    table.cell(7, 5);
    verify(table.firstColumn() == 3);
    verify(table.firstRow() == 5);
    verify(table.lastColumn() == 7);
    verify(table.lastRow() == 15);

    // getName() & setName()
    verify(_tcscmp(table.getName(), _T("table")) == 0);
    table.setName(_T("table"));
    verify(_tcscmp(table.getName(), _T("table")) == 0);
    table.setName(_T("table other"));
    verify(_tcscmp(table.getName(), _T("table other")) == 0);

    // Error handling
    table.cell(0, 0);
    table.cell(255, 65535);
    try {
        table.cell(-1, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.cell(256, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.cell(0, -1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.cell(0, 65536);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        table.cell(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.cell(_T("A"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        table.clearCell(-1, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearCell(0, -1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        table.clearCell(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearCell(_T("A"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        table.clearRange(-1, 0, 0, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearRange(0, -1, 0, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearRange(0, 0, -1, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearRange(0, 0, 0, -1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        table.clearRange(0, _T("A1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearRange(_T("A1"), 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearRange(_T("A"), _T("A1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.clearRange(_T("A1"), _T("A"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        table.isEmptyCell(-1, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.isEmptyCell(0, -1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        table.isEmptyCell(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table.isEmptyCell(_T("A"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    while (sc.tableCount()) {
        sc.removeTable(0);
    }
    sc.insertTable(0, _T("table 1"));
    splib::Table& table2 = sc.insertTable(0, _T("table 2"));
    try {
        table2.setName(_T("table 1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    sc.insertTable(0, _T("table 3"));
    try {
        table2.setName(_T("table 1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table2.setName(_T("table 3"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    table2.setName(_T("table 2"));
    table2.setName(_T("table 4"));

    // various malformed cell specs
    try {
        table2.cell(_T("B"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table2.cell(_T("1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table2.cell(_T("B1A"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table2.cell(_T("B1."));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table2.cell(_T(".B1"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        table2.cell(_T("B0"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    verify(table2.isEmptyCell(_T("IV10")));
}

void testCell() {
    splib::SpreadsheetImpl sc;
    splib::Table& table = sc.insertTable(0, _T("table"));
    splib::Cell& cell = table.cell(_T("C3"));

    // default state
    verify(cell.getType() == splib::Cell::NONE);
    verify(cell.getHAlignment() == splib::Cell::HADEFAULT);
    verify(cell.getVAlignment() == splib::Cell::VADEFAULT);

    // TEXT
    cell.setText(_T("my text"));
    verify(cell.getType() == splib::Cell::TEXT);
    verify(_tcscmp(cell.getText(), _T("my text")) == 0);
    cell.setText(_T(""));
    verify(_tcscmp(cell.getText(), _T("")) == 0);
    cell.setText(0);
    verify(_tcscmp(cell.getText(), _T("")) == 0);

    // LONG
    cell.setLong(100);
    verify(cell.getType() == splib::Cell::LONG);
    verify(cell.getLong() == 100);
    cell.setLong(1000);
    verify(cell.getLong() == 1000);
    cell.setLong(0);
    verify(cell.getLong() == 0);
    cell.setLong(-1);
    verify(cell.getLong() == -1);

    // DOUBLE
    cell.setDouble(.5);
    verify(cell.getType() == splib::Cell::DOUBLE);
    verify(cell.getDouble() == .5);
    cell.setDouble(1.5);
    verify(cell.getDouble() == 1.5);
    cell.setDouble(0);
    verify(cell.getDouble() == 0);
    cell.setDouble(-.5);
    verify(cell.getDouble() == -.5);

    // DATE
    cell.setDate(splib::Date(_T("2006-07-12")));
    verify(cell.getType() == splib::Cell::DATE);
    verify(cell.getDate() == splib::Date(_T("2006-07-12")));
    cell.setDate(splib::Date(_T("2006-07-15")));
    verify(cell.getDate() == splib::Date(_T("2006-07-15")));

    // TIME
    cell.setTime(splib::Time(_T("12:15")));
    verify(cell.getType() == splib::Cell::TIME);
    verify(cell.getTime() == splib::Time(_T("12:15")));
    cell.setTime(splib::Time(_T("21:30:45")));
    verify(cell.getTime() == splib::Time(_T("21:30:45")));

    // FORMULA
    cell.setFormula(_T("SUM(A1:B10)"));
    verify(cell.getType() == splib::Cell::FORMULA);
    verify(_tcscmp(cell.getFormula(), _T("SUM(A1:B10)")) == 0);
    cell.setFormula(_T("SUM(B1:C1)"));
    verify(_tcscmp(cell.getFormula(), _T("SUM(B1:C1)")) == 0);

    // clear()
    cell.clear();
    verify(cell.getType() == splib::Cell::NONE);

    // getType() is tested by all the above code

    // h alignment
    verify(cell.getHAlignment() == splib::Cell::HADEFAULT);
    cell.setHAlignment(splib::Cell::LEFT);
    verify(cell.getHAlignment() == splib::Cell::LEFT);

    // v alignment
    verify(cell.getVAlignment() == splib::Cell::VADEFAULT);
    cell.setVAlignment(splib::Cell::BOTTOM);
    verify(cell.getVAlignment() == splib::Cell::BOTTOM);

    // exceptions
    cell.clear();
    try {
        cell.getText();
        verify(false);
    } catch (splib::IllegalStateException&) {
    }
    try {
        cell.getLong();
        verify(false);
    } catch (splib::IllegalStateException&) {
    }
    try {
        cell.getDouble();
        verify(false);
    } catch (splib::IllegalStateException&) {
    }
    try {
        cell.getDate();
        verify(false);
    } catch (splib::IllegalStateException&) {
    }
    try {
        cell.getTime();
        verify(false);
    } catch (splib::IllegalStateException&) {
    }
    try {
        cell.getFormula();
        verify(false);
    } catch (splib::IllegalStateException&) {
    }
}

void testDate() {
    // constructors and assignment
    verify(splib::Date() == splib::Date(0, 1, 1));
    verify(splib::Date(2006, 12, 13) == splib::Date(2006, 12, 13));
    verify(splib::Date(_T("2006-11-12")) == splib::Date(2006, 11, 12));
    splib::Date date1 = splib::Date(2006, 01, 31);
    splib::Date date2(date1);
    splib::Date date3;
    date3 = date1;
    verify(date2 == date1);
    verify(date3 == date1);

    // comparison
    verify(splib::Date(_T("2006-01-02")) == splib::Date(2006, 1, 2));
    verify(splib::Date(_T("2007-01-02")) != splib::Date(2006, 1, 2));
    verify(splib::Date(_T("2006-02-02")) != splib::Date(2006, 1, 2));
    verify(splib::Date(_T("2006-01-03")) != splib::Date(2006, 1, 2));

    // getters and setters
    splib::Date date = splib::Date(2006, 2, 28);
    verify(date.getYear() == 2006);
    date.setYear(2000);
    verify(date.getYear() == 2000);
    verify(date.getMonth() == 2);
    date.setMonth(6);
    verify(date.getMonth() == 6);
    verify(date.getDay() == 28);
    date.setDay(30);
    verify(date.getDay() == 30);
    date = splib::Date(2006, 2, 28);
    verify(_tcscmp(date.getString(), _T("2006-02-28")) == 0);
    date.setString(_T("1978-02-01"));
    verify(_tcscmp(date.getString(), _T("1978-02-01")) == 0);
    verify(_tcscmp(splib::Date(0, 1, 1).getString(), _T("0000-01-01")) == 0);

    // exceptions
    try {
        splib::Date(-1, 1, 1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Date(0, 1, 1);
    splib::Date(9999, 1, 1);
    try {
        splib::Date(10000, 1, 1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        splib::Date(2000, 0, 1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Date(2000, 1, 1);
    splib::Date(2000, 12, 1);
    try {
        splib::Date(2000, 13, 1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        splib::Date(2000, 1, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Date(2000, 1, 1);
    splib::Date(2000, 1, 31);
    try {
        splib::Date(2000, 1, 32);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    splib::Date(2000, 2, 29);
    try {
        splib::Date(2000, 2, 30);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Date(2001, 2, 28);
    try {
        splib::Date(2001, 2, 29);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Date(2004, 2, 29);
    try {
        splib::Date(2004, 2, 30);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Date(3000, 2, 28);
    try {
        splib::Date(3000, 2, 29);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    splib::Date(_T("0000-01-01"));
    try {
        splib::Date(_T("0000-00-01"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Date(_T("0000-01-00"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Date(_T("a0000-01-01"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Date(_T(""));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Date(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    date = splib::Date(_T("2006-07-12"));
    try {
        date.setYear(-1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        date.setMonth(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        date.setDay(32);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        date.setString(_T(""));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        date.setString(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
}

void testTime() {
    // constructors and assignment
    verify(splib::Time() == splib::Time(0, 0));
    verify(splib::Time(1, 2, 3, 4) == splib::Time(1, 2, 3, 4));
    verify(splib::Time(_T("02:03:04.567")) == splib::Time(2, 3, 4, 567));
    splib::Time time1 = splib::Time(01, 45);
    splib::Time time2(time1);
    splib::Time time3;
    time3 = time1;
    verify(time2 == time1);
    verify(time3 == time1);

    // comparison
    verify(splib::Time(_T("01:02")) == splib::Time(1, 2));
    verify(splib::Time(_T("01:02:00")) == splib::Time(1, 2));
    verify(splib::Time(_T("01:02:00.000")) == splib::Time(1, 2));
    verify(splib::Time(_T("02:02:00.000")) != splib::Time(1, 2));
    verify(splib::Time(_T("01:03:00.000")) != splib::Time(1, 2));
    verify(splib::Time(_T("01:02:01.000")) != splib::Time(1, 2));
    verify(splib::Time(_T("01:02:00.001")) != splib::Time(1, 2));

    // getters and setters
    splib::Time time = splib::Time(_T("13:45:35.111"));
    verify(time.getHours() == 13);
    time.setHours(15);
    verify(time.getHours() == 15);
    verify(time.getMinutes() == 45);
    time.setMinutes(30);
    verify(time.getMinutes() == 30);
    verify(time.getSeconds() == 35);
    time.setSeconds(51);
    verify(time.getSeconds() == 51);
    verify(time.getMillis() == 111);
    time.setMillis(998);
    verify(time.getMillis() == 998);
    time = splib::Time(_T("13:45:35.111"));
    verify(_tcscmp(time.getString(), _T("13:45:35.111")) == 0);
    time.setString(_T("12:15"));
    verify(_tcscmp(time.getString(), _T("12:15:00.000")) == 0);

    // parsing
    splib::Time(_T("13:45"));
    splib::Time(_T("13:45:15"));
    splib::Time(_T("13:45:15.123"));
    try {
        splib::Time(_T("13:45a"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T("13:45:a"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T("13:45:15:"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T("13:45:15.a"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T("13:45:15.123a"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    // exceptions
    try {
        splib::Time(-1, 0, 0, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Time(0, 0, 0, 0);
    splib::Time(23, 0, 0, 0);
    try {
        splib::Time(24, 0, 0, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        splib::Time(0, -1, 0, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Time(0, 0, 0, 0);
    splib::Time(0, 59, 0, 0);
    try {
        splib::Time(0, 60, 0, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        splib::Time(0, 0, -1, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Time(0, 0, 0, 0);
    splib::Time(0, 0, 59, 0);
    try {
        splib::Time(0, 0, 60, 0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    try {
        splib::Time(0, 0, 0, -1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    splib::Time(0, 0, 0, 0);
    splib::Time(0, 0, 0, 999);
    try {
        splib::Time(0, 0, 0, 1000);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    splib::Time(_T("00:00"));
    try {
        splib::Time(_T("24:00"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T("00:60"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T("00:00:60"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T("a00:00"));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(_T(""));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        splib::Time(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }

    time = splib::Time(_T("00:00"));
    try {
        time.setHours(-1);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        time.setMinutes(60);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        time.setSeconds(60);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        time.setMillis(1000);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        time.setString(_T(""));
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
    try {
        time.setString(0);
        verify(false);
    } catch (splib::IllegalArgumentException&) {
    }
}

void testWriters() {
    splib::SpreadsheetImpl sc;
    setupTestSpreadsheet(sc);
    splib::XlsWriter().write(sc, _T("testout.xls"));
    splib::XlsxWriter().write(sc, _T("testout.xlsx"));
    splib::OdsWriter().write(sc, _T("testout.ods"));
}

void setupTestSpreadsheet(splib::Spreadsheet& sc) {
    setupCellTypesTable(sc);
    setupFormulasTable(sc);
    setupAlignmentTable(sc);
    setupRowsAndColumnsTable(sc);    
    setupSpecialCharsTable(sc);
    setupHeightsAndWidthsTable(sc);
    setupEmptyTable(sc);
}

void setupCellTypesTable(splib::Spreadsheet& sc) {
    splib::Table& table = sc.insertTable(sc.tableCount(), _T("Cell Types"));
    int row = 0;
    table.cell(0, row).setText(_T("TEXT:"));
    for (int column = 1; column <= 5; column++) {
        std::basic_stringstream<_TCHAR> text;
        text << _T("text ") << column;
        table.cell(column, row).setText(text.str().c_str());
    }
    table.cell(0, ++row).setText(_T("LONG:"));
    for (int column = 1; column <= 5; column++) {
        table.cell(column, row).setLong(column);
    }
    table.cell(0, ++row).setText(_T("DOUBLE:"));
    for (int column = 1; column <= 5; column++) {
        table.cell(column, row).setDouble(1. / column);
    }
    table.cell(0, ++row).setText(_T("DATE:"));
    for (int column = 1; column <= 5; column++) {
        table.cell(column, row).setDate(
            splib::Date(2006 + column, 1 + column, 12 + column));
    }
    table.cell(0, ++row).setText(_T("TIME:"));
    for (int column = 1; column <= 5; column++) {
        table.cell(column, row).setTime(
            splib::Time(12 + column, 30 + column, 0 + column));
    }
    for (int column = 1; column <= 5; column++) {
        table.columns().get(column).setWidth(70);
    }
}

void setupFormulasTable(splib::Spreadsheet& sc) {
    splib::Table& table = sc.insertTable(sc.tableCount(), _T("Formulas"));
    const int W = 5;  // width of the table
    const int H = 10; // height of the table
    for (int c = 0; c < W; c++) {
        for (int r = 0; r < H; r++) {
            table.cell(c, r).setLong(c * H + r);
        }
        std::basic_stringstream<_TCHAR> formula;
        formula << _T("SUM(A1:") << (_TCHAR)('A' + c) << H << _T(")");
        table.cell(c, H).setFormula(formula.str().c_str());
    }
}

void setupAlignmentTable(splib::Spreadsheet& sc) {
    splib::Table& table = sc.insertTable(sc.tableCount(), _T("Alignment"));
    const int W = 10; // number of alignement combinations to demo
    int row = 0;
    table.cell(0, row).setText(_T("TEXT:"));
    for (int column = 1; column <= W; column++) {
        std::basic_stringstream<_TCHAR> text;
        text << _T("text ") << column;
        table.cell(column, row).setText(text.str().c_str());
    }
    table.cell(0, ++row).setText(_T("LONG:"));
    for (int column = 1; column <= W; column++) {
        table.cell(column, row).setLong(column);
    }
    table.cell(0, ++row).setText(_T("DOUBLE:"));
    for (int column = 1; column <= W; column++) {
        table.cell(column, row).setDouble(1. / column);
    }
    table.cell(0, ++row).setText(_T("DATE:"));
    for (int column = 1; column <= W; column++) {
        table.cell(column, row).setDate(
            splib::Date(2006 + column, 1 + column, 12 + column));
    }
    table.cell(0, ++row).setText(_T("TIME:"));
    for (int column = 1; column <= W; column++) {
        table.cell(column, row).setTime(
            splib::Time(12 + column, 30 + column, 0 + column));
    }
    for (int row = 0; row < 5; row++) {
        for (int column = 1; column <= W; column++) {
            splib::Cell::HAlignment hAlign = splib::Cell::HADEFAULT;
            splib::Cell::VAlignment vAlign = splib::Cell::VADEFAULT;
            switch (column) {
                case 1:
                    hAlign = splib::Cell::HADEFAULT;
                    vAlign = splib::Cell::VADEFAULT;
                    break;
                case 2:
                    hAlign = splib::Cell::CENTER;
                    vAlign = splib::Cell::MIDDLE;
                    break;
                case 3:
                    hAlign = splib::Cell::CENTER;
                    vAlign = splib::Cell::TOP;
                    break;
                case 4:
                    hAlign = splib::Cell::RIGHT;
                    vAlign = splib::Cell::TOP;
                    break;
                case 5:
                    hAlign = splib::Cell::RIGHT;
                    vAlign = splib::Cell::MIDDLE;
                    break;
                case 6:
                    hAlign = splib::Cell::RIGHT;
                    vAlign = splib::Cell::BOTTOM;
                    break;
                case 7:
                    hAlign = splib::Cell::CENTER;
                    vAlign = splib::Cell::BOTTOM;
                    break;
                case 8:
                    hAlign = splib::Cell::LEFT;
                    vAlign = splib::Cell::BOTTOM;
                    break;
                case 9:
                    hAlign = splib::Cell::LEFT;
                    vAlign = splib::Cell::MIDDLE;
                    break;
                case 10:
                    hAlign = splib::Cell::LEFT;
                    vAlign = splib::Cell::TOP;
                    break;
            }
            table.cell(column, row).setHAlignment(hAlign);
            table.cell(column, row).setVAlignment(vAlign);
        }
    }

    for (int column = 1; column <= W; column++) {
        table.columns().get(column).setWidth(80);
    }
    for (int row = 0; row < 5; row++) {
        table.rows().get(row).setHeight(30);
    }
}

void setupRowsAndColumnsTable(splib::Spreadsheet& sc) {
    splib::Table& table = sc.insertTable(sc.tableCount(), _T("Rows & Columns"));
    const int BIG_SQUARE_SIZE = 5;
    const int SMALL_SQUARE_SIZE = 5;
    const int SMALL_SQUARE_MARGIN = 2;
    const int OFF = 3;
    for (int y = 0; y < BIG_SQUARE_SIZE; y++) {
        for (int x = 0; x < BIG_SQUARE_SIZE; x++) {
            int startRow = (SMALL_SQUARE_SIZE + SMALL_SQUARE_MARGIN) * y + OFF;
            int startCol = (SMALL_SQUARE_SIZE + SMALL_SQUARE_MARGIN) * x + OFF;
            for (int row = 0; row < SMALL_SQUARE_SIZE; row++) {
                for (int col = 0; col < SMALL_SQUARE_SIZE; col++) {
                    int value = row * SMALL_SQUARE_SIZE + col;
                    table.cell(startCol + col, startRow + row).setLong(value);
                }
            }
        }
    }
}

void setupSpecialCharsTable(splib::Spreadsheet& sc) {
    // warning: for unknown reason excel does not allow single quote at
    // the end of the sheet name
    int index = sc.tableCount();
    splib::Table& table = sc.insertTable(index, _T("Special Chars &<>\"' "));
    table.cell(0, 0).setText(_T("ampersand: &&&&&"));
    table.cell(0, 1).setText(_T("less than: <<<<<"));
    table.cell(0, 2).setText(_T("greater than: >>>>>"));
    table.cell(0, 3).setText(_T("double quote: \"\"\"\"\""));
    table.cell(0, 4).setText(_T("single quote: '''''"));
    table.columns().get(_T("A")).setWidth(150);
}

void setupHeightsAndWidthsTable(splib::Spreadsheet& sc) {
    int index = sc.tableCount();
    splib::Table& table = sc.insertTable(index, _T("Heights & Widths"));
    for (int row = 0; row < 10; row++) {
        table.rows().get(row + 3).setHeight(5 * pow(2., row / 2.));
    }
    for (int column = 0; column < 10; column++) {
        table.columns().get(column + 2).setWidth(10 * pow(2., column / 2.));
    }
}

void setupEmptyTable(splib::Spreadsheet& sc) {
    sc.insertTable(sc.tableCount(), _T("Empty Sheet"));
}
