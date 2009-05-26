// File: TableImpl.h
// TableImpl declaration file
//

#ifndef TABLEIMPL_H
#define TABLEIMPL_H

#include "splib.h"
#include "RowsImpl.h"
#include "ColumnsImpl.h"

namespace splib {

/** Default implementation of the <code>Table</code> interface. */
class TableImpl : public Table {
    public:
        /**
         * Creates a new instance of <code>TableImpl</code>.
         * @param spreadsheet a pointer to the parent spreadsheet;
         *        used to protect against table name duplications
         */
        TableImpl(Spreadsheet* spreadsheet);

        /** Destructor */
        virtual ~TableImpl();

        // inherit doc
        virtual Cell& cell(int column, int row);

        // inherit doc
        virtual Cell& cell(const _TCHAR* columnrow);

        // inherit doc
        virtual void clearCell(int column, int row);

        // inherit doc
        virtual void clearCell(const _TCHAR* columnrow);

        // inherit doc
        virtual void clearRange(int column1, int row1,
                                int column2, int row2);

        // inherit doc
        virtual void clearRange(const _TCHAR* columnrow1,
                                const _TCHAR* columnrow2);

        // inherit doc
        virtual bool isEmptyCell(int column, int row) const;

        // inherit doc
        virtual bool isEmptyCell(const _TCHAR* columnrow) const;

        // inherit doc
        virtual int firstColumn() const;

        // inherit doc
        virtual int lastColumn() const;

        // inherit doc
        virtual int firstRow() const;

        // inherit doc
        virtual int lastRow() const;

        // inherit doc
        virtual Rows& rows();

        // inherit doc
        virtual Columns& columns();

        // inherit doc
        virtual const _TCHAR* getName() const;

        // inherit doc
        virtual void setName(const _TCHAR* name);

    private:
        /** The collection of rows */
        RowsImpl rowsImpl;

        /** The collection of columns */
        ColumnsImpl columnsImpl;

        /** The table name */
        std::basic_string<_TCHAR> name;

        /** A pointer to the parent spreadsheet */
        Spreadsheet* spreadsheet;
};

}

#endif // TABLEIMPL_H
