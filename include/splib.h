// File: splib.h
// Public splib declarations. The project that
// uses splib usually includes this file.
//

#ifndef SPLIB_H
#define SPLIB_H

#include <vector>
#include <string>

// platform-specific includes and declarations
#ifdef WIN32
    // this symbol should not be defined on the Win32 project that uses
    // splib.dll; this way the project will see SPLIB_API functions as being
    // imported from a DLL
    #ifdef SPLIB_EXPORTS
        #define SPLIB_API __declspec(dllexport)
    #else
        #define SPLIB_API __declspec(dllimport)
    #endif // SPLIB_EXPORTS
    #include <tchar.h>
#else
    #define SPLIB_API
    typedef char _TCHAR;
#endif // WIN32

namespace splib {

// forward declarations
class Table;

/**
 * A spreadsheet object. The spreadsheet consists of a number of tables.
 * This interface provides methods for adding, removing, and accessing
 * the tables in the spreadsheet.
 */
class SPLIB_API Spreadsheet {
    public:
        /**
         * Creates a new table and inserts it into the spreadsheet.
         * @param index the index at which to insert the table
         * @param name a pointer to the name of the table
         * @return a reference to the newly created table
         * @throw IllegalArgumentException if <code>index</code> is negative
         *        or greater than the value returned by tableCount();
         *        or if <code>name</code> is 0;
         *        or if a table with the specified name already exists
         */
        virtual Table& insertTable(int index, const _TCHAR* name) = 0;

        /**
         * Retrieves a table by index.
         * @param index the index at which to retrieve the table
         * @return a reference to the table object
         * @throw IllegalArgumentException if <code>index</code> is
         *        negative or greater than or equal to the value returned by
         *        tableCount()
         */
        virtual Table& table(int index) = 0;

        /**
         * Retrieves a table by name.
         * @param name the name of the table to retrieve
         * @return a reference to the table object
         * @throw IllegalArgumentException if <code>name</code> is 0
         *        or if a table with the specified name does not exist
         *        in the spreadsheet
         */
        virtual Table& table(const _TCHAR* name) = 0;

        /**
         * Retrieves the number of tables in the spreadsheet.
         * @return the number of tables in the spreadsheet
         */
        virtual int tableCount() const = 0;

        /**
         * Determines if a table with the specified name already exists in
         * the spreadsheet.
         * @param name a pointer to a string that specifies the table name
         * @return true if a table with the specified name exists;
         *         false otherwise
         * @throw IllegalArgumentException if <code>name</code> is 0
         */
        virtual bool tableExists(const _TCHAR* name) = 0;

        /**
         * Removes a table from the spreadsheet.
         * @param index the index at which to remove the table
         * @throw IllegalArgumentException if <code>index</code> is
         *        negative or greater than or equal to the value returned by
         *        tableCount()
         */
        virtual void removeTable(int index) = 0;

        /**
         * Removes a table from the spreadsheet.
         * @param name the name of the table to remove
         * @throw IllegalArgumentException if <code>name</code> is 0
         *        or if a table with the specified name does not exist
         *        in the spreadsheet
         */
        virtual void removeTable(const _TCHAR* name) = 0;

        /** Empty virtual destructor */
        virtual ~Spreadsheet() {}
};

// forward declarations
class Cell;
class Rows;
class Columns;

/**
 * A spreadsheet table. This interface provides methods for accessing
 * table cells as well as other table properties.
 */
class SPLIB_API Table {
    public:
        /**
         * Retrieves a cell at the specified position.
         * If the cell at the specified position is empty (that is, no cell
         * object exists for it), the corresponding cell object is created
         * and initialized to default, and a reference to that object is
         * returned. If the cell at the specified position is not empty
         * (that is, a cell object exists for it), a reference to that object
         * is returned.
         * @param column column index for which to retrieve the cell
         * @param row row index for which to retrieve the cell
         * @return a reference to the cell object
         * @throw IllegalArgumentException if <code>column</code> or
         *        <code>row</code> is invalid
         */
        virtual Cell& cell(int column, int row) = 0;

        /**
         * Retrieves a cell at the specified position.
         * The position is specified in the
         * &lt;column&nbsp;letter(s)&gt;&lt;row&nbsp;number&gt; format used
         * by most spreadsheet applications to identify cells.
         * If the cell at the specified position is empty (that is, no cell
         * object exists for it), the corresponding cell object is created
         * and initialized to default, and a reference to that object is
         * returned. If the cell at the specified position is not empty
         * (that is, a cell object exists for it), a reference to that object
         * is returned.
         * @param columnrow a string in the
         *        &lt;column&nbsp;letter(s)&gt;&lt;row&nbsp;number&gt;
         *        format; for example, <code>"A1"</code>, <code>"D20"</code>,
         *        <code>"AA9"</code>
         * @return a reference to the cell object
         * @throw IllegalArgumentException if <code>columnrow</code> is invalid
         */
        virtual Cell& cell(const _TCHAR* columnrow) = 0;

        /**
         * Clears the cell at the specified position. Deletes the cell
         * object associated with the position, if it exists.
         * @param column column index
         * @param row row index
         * @throw IllegalArgumentException if <code>column</code> or
         *        <code>row</code> is invalid
         */
        virtual void clearCell(int column, int row) = 0;

        /**
         * Clears the cell at the specified position. Deletes the cell
         * object associated with the position, if it exists.
         * @param columnrow a string in the
         *        &lt;column&nbsp;letter(s)&gt;&lt;row&nbsp;number&gt; format
         * @throw IllegalArgumentException if <code>columnrow</code> is invalid
         */
        virtual void clearCell(const _TCHAR* columnrow) = 0;

        /**
         * Clears a rectangular range of cells. The range is identified by
         * the positions of the two corners of the rectangle. The rectangle
         * does not have to be normalized.
         * @param column1 column index of the top left corner of the range
         * @param row1 row index of the top left corner of the range
         * @param column2 column index of the bottom right corner of the range
         * @param row2 row index of the bottom right corner of the range
         * @throw IllegalArgumentException if <code>column1</code>,
         *        <code>row1</code>, <code>column2</code>, or
         *        <code>row2</code> is invalid
         */
        virtual void clearRange(int column1, int row1,
                                int column2, int row2) = 0;

        /**
         * Clears a rectangular range of cells. The range is identified by
         * the positions of the two corners of the rectangle. The rectangle
         * does not have to be normalized.
         * @param columnrow1 a string in the
         *        &lt;column&nbsp;letter(s)&gt;&lt;row&nbsp;number&gt; format
         *        that identifies the top left corner of the range
         * @param columnrow2 a string in the
         *        &lt;column&nbsp;letter(s)&gt;&lt;row&nbsp;number&gt; format
         *        that identifies the bottom right corner of the range
         * @throw IllegalArgumentException if <code>columnrow1</code> or
         *        <code>columnrow2</code> is invalid
         */
        virtual void clearRange(const _TCHAR* columnrow1,
                                const _TCHAR* columnrow2) = 0;

        /**
         * Determines if the specified position is empty.
         * @param column column index
         * @param row row index
         * @return true if the specified position is empty (that is, no
         *         cell object exists at it); false otherwise
         * @throw IllegalArgumentException if <code>column</code> or
         *        <code>row</code> is invalid
         */
        virtual bool isEmptyCell(int column, int row) const = 0;

        /**
         * Determines if the specified position is empty.
         * @param columnrow a string in the
         *        &lt;column&nbsp;letter(s)&gt;&lt;row&nbsp;number&gt; format
         * @return true if the specified position is empty (that is, no
         *         cell object exists at it); false otherwise
         * @throw IllegalArgumentException if <code>columnrow</code> is
         *        invalid
         */
        virtual bool isEmptyCell(const _TCHAR* columnrow) const = 0;

        /**
         * Determines the first (leftmost) table column containing
         * non-empty cells.
         * @return the index of the leftmost table column containing
         *         non-empty cells. -1 if there are no cells in the table.
         */
        virtual int firstColumn() const = 0;

        /**
         * Determines the first (topmost) table row containing
         * non-empty cells.
         * @return the index of the topmost table row containing
         *         non-empty cells. -1 if there are no cells in the table.
         */
        virtual int firstRow() const = 0;

        /**
         * Determines the last (rightmost) table column containing
         * non-empty cells.
         * @return the index of the rightmost table column containing
         *         non-empty cells. -1 if there are no cells in the table.
         */
        virtual int lastColumn() const = 0;

        /**
         * Determines the last (bottommost) table row containing
         * non-empty cells.
         * @return the index of the bottommost table row containing
         *         non-empty cells. -1 if there are no cells in the table.
         */
        virtual int lastRow() const = 0;

        /**
         * Retrieves the collection of rows.
         * @return collection of rows.
         */
        virtual Rows& rows() = 0;

        /**
         * Retrieves the collection of columns.
         * @return collection of columns.
         */
        virtual Columns& columns() = 0;

        /**
         * Retrieves the name of the table.
         * @return a pointer to the string containing the name of the table
         */
        virtual const _TCHAR* getName() const = 0;

        /**
         * Sets the name of the table.
         * @param name a pointer to the string specifying the new name
         *        of the table
         * @throw IllegalArgumentException if a table with this name
         *        already exists in the spreadsheet
         */
        virtual void setName(const _TCHAR* name) = 0;

        /** Empty virtual destructor */
        virtual ~Table() {}
};

// forward declarations
class Date;
class Time;

/**
 * A spreadsheet cell. This interface defines methods to obtain and modify
 * cell properties such as cell content, type, and formatting.
 */
class SPLIB_API Cell {
    public:
        /**
         * Retrieves the cell text.
         * The method succeeds if the current cell type is
         * <code>TEXT</code>, otherwise IllegalStateException
         * is thrown.
         * @return a pointer to the string containing the cell text
         * @throw IllegalStateException if the current cell type is not
         *        <code>TEXT</code>
         */
        virtual const _TCHAR* getText() const = 0;

        /**
         * Sets the cell type to <code>TEXT</code> and updates the cell
         * text.
         * @param text a pointer to the string specifying the new cell text;
         *        <code>0</code> is equivalent to <code>""</code>
         */
        virtual void setText(const _TCHAR* text) = 0;

        /**
         * Retrieves the long integer value of the cell.
         * The method succeeds if the current cell type is
         * <code>LONG</code>, otherwise IllegalStateException
         * is thrown.
         * @return the value of the cell
         * @throw IllegalStateException if the current cell type is not
         *        <code>LONG</code>
         */
        virtual long getLong() const = 0;

        /**
         * Sets the cell type to <code>LONG</code> and updates the value
         * of the cell.
         * @param l new long integer value of the cell
         */
        virtual void setLong(long l) = 0;

        /**
         * Retrieves the double value of the cell.
         * The method succeeds if the current cell type is
         * <code>DOUBLE</code>, otherwise IllegalStateException
         * is thrown.
         * @return the value of the cell
         * @throw IllegalStateException if the current cell type is not
         *        <code>DOUBLE</code>
         */
        virtual double getDouble() const = 0;

        /**
         * Sets the cell type to <code>DOUBLE</code> and updates the value
         * of the cell.
         * @param d new value of the cell
         */
        virtual void setDouble(double d) = 0;

        /**
         * Retrieves the date stored in the cell.
         * The method succeeds if the current cell type is
         * <code>DATE</code>, otherwise IllegalStateException
         * is thrown.
         * @return the date stored in the cell
         * @throw IllegalStateException if the current cell type is not
         *        <code>DATE</code>
         */
        virtual const Date& getDate() const = 0;

        /**
         * Sets the cell type to <code>DATE</code> and updates the date
         * stored in the cell.
         * @param date the new date to store in the cell
         */
        virtual void setDate(const Date& date) = 0;

        /**
         * Retrieves the time stored in the cell.
         * The method succeeds if the current cell type is
         * <code>TIME</code>, otherwise IllegalStateException
         * is thrown.
         * @return the time stored in the cell
         * @throw IllegalStateException if the current cell type is not
         *        <code>DATE</code>
         */
        virtual const Time& getTime() const = 0;

        /**
         * Sets the cell type to <code>TIME</code> and updates the time
         * stored in the cell.
         * @param time the new time to store in the cell
         */
        virtual void setTime(const Time& time) = 0;

        /**
         * Retrieves the formula stored in the cell.
         * The method succeeds if the current cell type is
         * <code>FORMULA</code>, otherwise IllegalStateException
         * is thrown.
         * @return the formula stored in the cell
         * @throw IllegalStateException if the current cell type is not
         *        <code>FORMULA</code>
         */
        virtual const _TCHAR* getFormula() const = 0;

        /**
         * Sets the cell type to <code>FORMULA</code> and updates the
         * formula stored in the cell.
         * Only sums of cell ranges are supported, such as "SUM(A1:B10)".
         * @param formula the new formula to store in the cell
         * @throw IllegalArgumentException if <code>formula</code>
         *        is invalid
         */
        virtual void setFormula(const _TCHAR* formula) = 0;

        /**
         * Sets the cell type to <code>NONE</code> and clears all data
         * stored in the cell.
         */
        virtual void clear() = 0;

        /** Possible cell types. */
        enum Type {
            /** No type (empty cell) */
            NONE = 0, 
            /** Text */
            TEXT, 
            /** Long integer */
            LONG,
            /** Double precision floating point number  */
            DOUBLE,
            /** Date */
            DATE, 
            /** Time */
            TIME,
            /** Formula */
            FORMULA};

        /**
         * Retrieves the cell type.
         * @return the cell type
         */
        virtual Type getType() const = 0;

        /**
         * The enumeration lists all possible horizontal alignment
         * possibilities.
         */
        enum HAlignment {
            /** Default alignment */
            HADEFAULT = 0,
            /** Left alignment */
            LEFT,
            /** Center alignment */
            CENTER,
            /** Right alignment */
            RIGHT,
            /** Justified alignment */
            JUSTIFIED,
            /** Filled alignment */
            FILLED
        };

        /**
         * Retrieves the horizontal alignment of the cell.
         * @return horizontal alignment
         */
        virtual HAlignment getHAlignment() const = 0;

        /**
         * Sets the horizontal alignment of the cell.
         * @param hAlignment horizontal alignment
         */
        virtual void setHAlignment(HAlignment hAlignment) = 0;

        /**
         * The enumeration lists all possible vertical alignment
         * possibilities.
         */
        enum VAlignment {
            /** Default alignment */
            VADEFAULT = 0, 
            /** Top alignment */
            TOP,
            /** Middle alignment */
            MIDDLE,
            /** Bottom alignment */
            BOTTOM
        };

        /**
         * Retrieves the vertical alignment of the cell.
         * @return vertical alignment
         */
        virtual VAlignment getVAlignment() const = 0;

        /**
         * Sets the vertical alignment of the cell.
         * @param vAlignment vertical alignment
         */
        virtual void setVAlignment(VAlignment vAlignment) = 0;

        /** Empty virtual destructor */
        virtual ~Cell() {}
};

/**
 * An indexed collection of objects. Each object is accessible by an integer
 * index. The objects are kept sorted by their indices.
 */
template<class T>
class IndexedCollection {
    public:
        /**
         * This class represents an entry stored in the collection.
         * Each entry references an object and an index of that object.
         */
        class Entry {
            public:
                /**
                 * Creates a new <code>Entry</code>.
                 * @param object reference to the object 
                 * @param index index of the object
                 */
                Entry(T& object, int index) : obj(object), ind(index) {}

                /**
                 * Copy constructor.
                 * @param entry existing entry from which to construct
                 *        a new entry 
                 */
                Entry(const Entry& entry) : obj(entry.obj), ind(entry.ind) {}

                /**
                 * Retrieves the object referenced by this entry.
                 * @return a reference to the object referenced by this entry
                 */
                T& object() {return this->obj;}

                /**
                 * Retrieves the index of the object referenced by this entry.
                 * @return index of the object referenced by this entry
                 */
                int index() {return this->ind;}
                
            private:
                /**
                 * Assignment operator. Declared private to disallow
                 * assignments.
                 */
                Entry& operator = (const Entry& entry) {return *this;}

            private:
                /** Object referenced by this entry */
                T& obj;

                /** Index of the object */
                int ind;
        };

        /** An iterator over objects in the collection. */
        class Iterator {
            public:
                /**
                 * Returns true if the iteration has more elements.
                 * @return true if the iteration has more elements,
                 *         false otherwise
                 */
                virtual bool hasNext() const = 0;

                /**
                 * Returns the next element in the iteration.
                 * @return an Entry that contains a reference to the next
                 *         object and the index of that object
                 * @throw IllegalStateException if iteration has no more
                 *        elements
                 */
                virtual Entry next() = 0;
                
                /** Empty virtual destructor */
                virtual ~Iterator() {}
        };

    public:
        /**
         * Retrieves an object by index. If the object with the given
         * index does not exist, it is created.
         * @param index index of the object
         * @return a reference to the object
         */
        virtual T& get(int index) = 0;

        /**
         * Returns true if this collection contains the object with
         * the specified index.
         * @param index index whose presence in this collection is to
         *        be tested
         * @return true if this collection contains the specified index,
         *         false otherwise
         */
        virtual bool contains(int index) = 0;

        /**
         * Removes the object with the given index if it is present in this
         * collection.
         * @param index index whose object is to be removed
         */
        virtual void remove(int index) = 0;

        /**
         * Returns the number of objects in this collection.
         * @return the number of objects in this collection
         */
        virtual int size() const = 0;

        /**
         * Returns true if this collection contains no objects.
         * @return true if this collection contains no objects,
         *         false otherwise
         */
        virtual bool isEmpty() const = 0;

        /**
         * Returns an iterator over the objects in this collection.
         * The iterator returns objects in the order of their indices.
         * @return an iterator over the objects in this collection
         */
        virtual Iterator* iterator() = 0;

        /**
         * Returns the first object in the collection.
         * @return an Entry that contains a reference to the first
         *         object and the index of that object
         * @throw IllegalStateException if this collection contains no
         *        objects
         */
        virtual Entry first() = 0;
        
        /**
         * Returns the last object in the collection.
         * @return an Entry that contains a reference to the last
         *         object and the index of that object
         * @throw IllegalStateException if this collection contains no
         *        objects
         */
        virtual Entry last() = 0;

        /** Empty virtual destructor */
        virtual ~IndexedCollection() {}
};

/**
 * A collection of cells of a single row. The semantics of the
 * collection is defined by the <code>IndexedCollection</code> interface.
 */
class SPLIB_API Cells : public IndexedCollection<Cell> {
};

/** A row of a table. */
class SPLIB_API Row {
    public:
        /**
         * Returns the height of the row.
         * @return height of the row, in points. If negative,
         *         default row height is used.
         */
        virtual double getHeight() const = 0;

        /**
         * Sets the height of the row.
         * @param height new height of the row, in points. If negative,
         *        default row height is used.
         */
        virtual void setHeight(double height) = 0;

        /**
         * Returns the collection of cells of this row.
         * @return the collection of cells of this row
         */
        virtual Cells& cells() = 0;

        /** Empty virtual destructor */
        virtual ~Row() {}
};

/**
 * A collection of rows. The semantics of the collection is defined
 * by the <code>IndexedCollection</code> interface.
 */
class SPLIB_API Rows : public IndexedCollection<Row> {
};

/** A column of a table. */
class SPLIB_API Column {
    public:
        /**
         * Returns the width of the column.
         * @return width of the column, in points. If negative, default
         *         column width is used.
         */
        virtual double getWidth() const = 0;

        /**
         * Sets the width of the column.
         * @param width new width of the column, in points. If negative,
         *        default column width is used.
         */
        virtual void setWidth(double width) = 0;

        /** Empty virtual destructor */
        virtual ~Column() {}
};

/**
 * A collection of columns. The semantics of the collection is defined by
 * the <code>IndexedCollection</code> interface, plus this interface adds
 * the ability to retrieve a column by column string.
 */
class SPLIB_API Columns : public IndexedCollection<Column> {
    public:
        /**
         * Retrieves a column by index. If the column with the given
         * index does not exist, it is created.
         * @param index index of the column
         * @return reference to the column
         */
        virtual Column& get(int index) = 0;

        /**
         * Retrieves a column by column character (or character string).
         * If the specified column does not exist, it is created.
         * @param column string that identifies the column, for example, "AA"
         * @return reference to the column
         */
        virtual Column& get(const _TCHAR* column) = 0;
};

/**
 * This class represents a date. The date consists of three numbers:
 * <code>year</code>, <code>month</code>, and <code>day</code>. These
 * numbers always form a valid Gregorian calendar date, see e.g.
 * http://www.timeanddate.com/date/leapyear.html. Valid year range is
 * <code>0..9999</code>.
 */
class SPLIB_API Date {
    public:
        /**
         * Creates a new instance of <code>Date</code> and initializes
         * it to <code>0000-01-01</code>.
         */
        Date();

        /**
         * Creates a new instance of <code>Date</code> and initializes
         * it with specific numbers.
         * @param year year
         * @param month month
         * @param day day
         * @throw IllegalArgumentException if the numbers
         *        do not form a valid date
         */
        Date(int year, int month, int day);

        /**
         * Creates a new instance of <code>Date</code> and initializes
         * it with a string in the <code>YYYY-MM-DD</code> format.
         * @param date a pointer to a string in the <code>YYYY-MM-DD</code>
         *        format; examples: <code>"2006-07-12"</code>,
         *        <code>"2000-12-25"</code>
         * @throw IllegalArgumentException if <code>date</code>
         *        does not represent a valid date
         */
        Date(const _TCHAR* date);

        /** Destructor */
        virtual ~Date();

        /**
         * == operator
         * @param date an object to compare to
         */
        bool operator == (const Date& date) const;

        /**
         * != operator
         * @param date an object to compare to
         */
        bool operator != (const Date& date) const;

        /**
         * Retrieves the year.
         * @return year
         */
        int getYear() const;

        /**
         * Sets the year.
         * @param year year
         * @throw IllegalArgumentException if the new date
         *        would not be a valid date
         */
        void setYear(int year);

        /**
         * Retrieves the month.
         * @return month
         */
        int getMonth() const;

        /**
         * Sets the month.
         * @param month month
         * @throw IllegalArgumentException if the new date
         *        would not be a valid date
         */
        void setMonth(int month);

        /**
         * Retrieves the day.
         * @return day
         */
        int getDay() const;

        /**
         * Sets the day.
         * @param day day
         * @throw IllegalArgumentException if the new date
         *        would not be a valid date
         */
        void setDay(int day);

        /**
         * Retrieves the string representation of the date.
         * @return a pointer to a string in the <code>YYYY-MM-DD</code>
         *         format; examples: <code>"2006-07-12"</code>,
         *         <code>"2000-12-25"</code>
         */
        const _TCHAR* getString() const;

        /**
         * Initializes the date with a string in the <code>YYYY-MM-DD</code>
         * format.
         * @param date a pointer to a string in the <code>YYYY-MM-DD</code>
         *        format; examples: <code>"2006-07-12"</code>,
         *        <code>"2000-12-25"</code>
         * @throw IllegalArgumentException if <code>date</code>
         *        does not represent a valid date or is <code>0</code>
         */
        void setString(const _TCHAR* date);

    private:
        /** Validates a set of numbers. */
        bool validate(int year, int month, int day);

        /** Parses a string into numbers. */
        bool parse(const _TCHAR* date, int& year, int& month, int& day);

    private:
        /** The year */
        int year;

        /** The month */
        int month;

        /** The day */
        int day;

#pragma warning (disable: 4251)
        /** The string representation of the date. */
        std::basic_string<_TCHAR> str;
#pragma warning (default: 4251)
};

/**
 * This class represents a time value. The time consists of four numbers:
 * <code>hours</code>, <code>minutes</code>, <code>seconds</code>,
 * <code>milliseconds</code>. These numbers always form a valid
 * time for which the following conditions always hold:
 * <ul>
 * <li><code>0 <= hours <= 23</code>
 * <li><code>0 <= minutes <= 59</code>
 * <li><code>0 <= seconds <= 59</code>
 * <li><code>0 <= milliseconds <= 999</code>
 * </ul>
 */
class SPLIB_API Time {
    public:
        /**
         * Creates a new instance of <code>Time</code> and initializes
         * it to <code>00:00</code>.
         */
        Time();

        /**
         * Creates a new instance of <code>Time</code> and initializes
         * it with specific numbers.
         * @param hours hours
         * @param minutes minutes
         * @param seconds seconds
         * @param millis millis
         * @throw IllegalArgumentException if the numbers
         *        do not form a valid time
         */
        Time(int hours, int minutes, int seconds = 0, int millis = 0);

        /**
         * Creates a new instance of <code>Time</code> and initializes
         * it with a string in the <code>HH:MM[:SS[.mmm]]</code> format.
         * @param time a pointer to a string in the
         *        <code>HH:MM[:SS[.mmm]]</code> format;
         *        examples: <code>"23:59"</code>, <code>"23:59:59"</code>,
         *        <code>"23:59:59.999"</code>
         * @throw IllegalArgumentException if <code>time</code>
         *        does not represent a valid time
         */
        Time(const _TCHAR* time);

        /** Destructor */
        virtual ~Time();

        /**
         * == operator
         * @param time an object to compare to
         */
        bool operator == (const Time& time) const;

        /**
         * != operator
         * @param time an object to compare to
         */
        bool operator != (const Time& time) const;

        /**
         * Retrieves the hours.
         * @return hours
         */
        int getHours() const;

        /**
         * Sets the hours.
         * @param hours hours
         * @throw IllegalArgumentException if <code>hours</code> is not
         *        valid
         */
        void setHours(int hours);

        /**
         * Retrieves the minutes.
         * @return minutes
         */
        int getMinutes() const;

        /**
         * Sets the minutes.
         * @param minutes minutes
         * @throw IllegalArgumentException if <code>minutes</code> is not
         *        valid
         */
        void setMinutes(int minutes);

        /**
         * Retrieves the seconds.
         * @return seconds
         */
        int getSeconds() const;

        /**
         * Sets the seconds.
         * @param seconds seconds
         * @throw IllegalArgumentException if <code>seconds</code> is not
         *        valid
         */
        void setSeconds(int seconds);

        /**
         * Retrieves the milliseconds.
         * @return milliseconds
         */
        int getMillis() const;
        
        /**
         * Sets the milliseconds.
         * @param millis milliseconds
         * @throw IllegalArgumentException if <code>millis</code> is not
         *        valid
         */
        void setMillis(int millis);

        /**
         * Retrieves a string representation of the time.
         * @return a pointer to a string in the
         *        <code>HH:MM:SS.mmm</code> format;
         *        examples: <code>"23:59:00.000"</code>,
         *        <code>"23:59:59.000"</code>,
         *        <code>"23:59:59.999"</code>
         */
        const _TCHAR* getString() const;

        /**
         * Initializes the time with a string in the
         * <code>HH:MM[:SS[.mmm]]</code> format.
         * @param time a pointer to a string in the
         *        <code>HH:MM[:SS[.mmm]]</code> format;
         *        examples: <code>"23:59"</code>, <code>"23:59:59"</code>,
         *        <code>"23:59:59.999"</code>
         * @throw IllegalArgumentException if <code>time</code>
         *        does not represent a valid time or is <code>0</code>
         */
        void setString(const _TCHAR* time);

    private:
        /** Validates a set of numbers. */
        bool validate(int hours, int minutes, int seconds, int millis);

        /** Parses a string into numbers. */
        bool parse(const _TCHAR* time, int& hours, int& minutes, int& seconds,
                   int& millis);

    private:
        /** Hours */
        int hours;

        /** Minutes */
        int minutes;

        /** Seconds */
        int seconds;

        /** Milliseconds */
        int millis;

#pragma warning (disable: 4251)
        /** The string representation of the time. */
        std::basic_string<_TCHAR> str;
#pragma warning (default: 4251)
};

/** An object that can output a spreadsheet to a file. */
class SPLIB_API Writer {
    public:
        /**
         * Writes a spreadsheet to a file.
         * @param spreadsheet the spreadsheet to write
         * @param pathname a pointer to a 0-terminated path name
         *        to a file to write the spreadsheet to
         * @throws IOException if file creation or writing fails
         */
        virtual void write(Spreadsheet& spreadsheet,
                           const _TCHAR* pathname) = 0;

        /** Empty virtual destructor */
        virtual ~Writer() {}
};

/** Default implementation of the <code>Spreadsheet</code> interface. */
class SPLIB_API SpreadsheetImpl : public Spreadsheet {
    public:
        /** Creates a new instance of <code>SpreadsheetImpl</code>. */
        SpreadsheetImpl();

        /** Destructor */
        virtual ~SpreadsheetImpl();

        // inherit doc
        virtual Table& insertTable(int index, const _TCHAR* name);

        // inherit doc
        virtual Table& table(int index);

        // inherit doc
        virtual Table& table(const _TCHAR* name);

        // inherit doc
        virtual int tableCount() const;

        // inherit doc
        virtual bool tableExists(const _TCHAR* name);

        // inherit doc
        virtual void removeTable(int index);

        // inherit doc
        virtual void removeTable(const _TCHAR* name);

    private:
        /** Finds a table with the specified name and returns its index. */
        int find(const _TCHAR* name);

    private:
#pragma warning (disable: 4251)
        /** The collection of tables. */
        std::vector<Table*> tables;
#pragma warning (default: 4251)
};

/** Writer that outputs spreadsheets in Excel 97/2000 format. */
class SPLIB_API XlsWriter : public Writer {
    public:
        // inherit doc
        virtual void write(Spreadsheet& spreadsheet, const _TCHAR* pathname);
};

/** Writer that outputs spreadsheets in Excel 2007 format. */
class SPLIB_API XlsxWriter : public Writer {
    public:
        // inherit doc
        virtual void write(Spreadsheet& spreadsheet, const _TCHAR* pathname);
};

/** Writer that outputs spreadsheets in OpenDocument format. */
class SPLIB_API OdsWriter : public Writer {
    public:
        // inherit doc
        virtual void write(Spreadsheet& spreadsheet, const _TCHAR* pathname);
};

/**
 * An exception with an optional message. This interface declares
 * the method that can be used to retrieve the message.
 */
class SPLIB_API Exception {
    public:
        /**
         * Retrieves the message associated with the exception.
         * @return a pointer to the message string. May be 0.
         */
        virtual const _TCHAR* message() = 0;

        /** Empty virtual destructor */
        virtual ~Exception() {}
};

/** A reusable implementation of the <code>Exception</code> interface. */
class SPLIB_API ExceptionImpl : public Exception {
    public:
        /** Constructs a new exception with no message. */
        ExceptionImpl();

        /**
         * Constructs a new exception with a message.
         * @param message a pointer to the message string.
         */
        ExceptionImpl(const _TCHAR* message);

        // inherit doc
        virtual const _TCHAR* message();

    protected:
        /**
         * Sets the exception message. Can be called by a constructor
         * of the derived class.
         * @param message a pointer to the message string
         */
        void setMessage(const _TCHAR* message);

    private:
#pragma warning (disable: 4251)
        /** The object that stores the message string. */
        std::basic_string<_TCHAR> msg;
#pragma warning (default: 4251)
};

/**
 * An exception thrown to indicate that an argument passed to a method
 * is illegal.
 */
class SPLIB_API IllegalArgumentException : public ExceptionImpl {
    public:
        /** Constructs a new exception with no message. */
        IllegalArgumentException();

        /**
         * Constructs a new exception with a message.
         * @param message a pointer to the exception message
         */
        IllegalArgumentException(const _TCHAR* message);
};

/**
 * An exception thrown to indicate that a method has been invoked at
 * an illegal or inappropriate time.
 */
class SPLIB_API IllegalStateException : public ExceptionImpl {
    public:
        /** Constructs a new exception with no message. */
        IllegalStateException();

        /**
         * Constructs a new exception with a message.
         * @param message a pointer to the exception message
         */
        IllegalStateException(const _TCHAR* message);
};

/** An exception thrown to indicate an I/O error. */
class SPLIB_API IOException : public ExceptionImpl {
    public:
        /** Constructs a new exception with no message. */
        IOException();

        /**
         * Constructs a new exception with a message.
         * @param message a pointer to the exception message
         */
        IOException(const _TCHAR* message);
};

}

#endif // SPLIB_H
