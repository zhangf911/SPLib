// File: CellImpl.h
// CellImpl declaration file
//

#ifndef CELLIMPL_H
#define CELLIMPL_H

#include "splib.h"

namespace splib {

/** Default implementation of the <code>Cell</code> interface. */
class CellImpl : public Cell {
    public:
        /** Creates a new instance of <code>CellImpl</code>. */
        CellImpl();

        /** Destructor */
        virtual ~CellImpl();

        // inherit doc
        virtual const _TCHAR* getText() const;

        // inherit doc
        virtual void setText(const _TCHAR* text);

        // inherit doc
        virtual long getLong() const;

        // inherit doc
        virtual void setLong(long l);

        // inherit doc
        virtual double getDouble() const;

        // inherit doc
        virtual void setDouble(double d);

        // inherit doc
        virtual const Date& getDate() const;

        // inherit doc
        virtual void setDate(const Date& date);

        // inherit doc
        virtual const Time& getTime() const;

        // inherit doc
        virtual void setTime(const Time& time);

        // inherit doc
        virtual const _TCHAR* getFormula() const;

        // inherit doc
        virtual void setFormula(const _TCHAR* formula);

        // inherit doc
        virtual void clear();

        // inherit doc
        virtual Type getType() const;

        // inherit doc
        virtual HAlignment getHAlignment() const;

        // inherit doc
        virtual void setHAlignment(HAlignment hAlignment);

        // inherit doc
        virtual VAlignment getVAlignment() const;

        // inherit doc
        virtual void setVAlignment(VAlignment vAlignment);

    private:
        /** The cell type */
        Type type;

        /** The cell text */
        std::basic_string<_TCHAR> text;

        /** The set of possible simple values of the cell */
        union Value {
            long longValue;
            double doubleValue;
        };

        /** The union of simple values of the cell */
        Value value;

        /** The cell date */
        Date date;

        /** The cell time */
        Time time;

        /** The cell formula */
        std::basic_string<_TCHAR> formula;

        /** The horizontal alignment of the cell */
        HAlignment hAlignment;

        /** The vertical alignment of the cell */
        VAlignment vAlignment;
};

}

#endif // CELLIMPL_H
