// File: ColumnImpl.h
// ColumnImpl declaration file
//

#ifndef COLUMNIMPL_H
#define COLUMNIMPL_H

#include "splib.h"

namespace splib {

/** Default implementation of the <code>Column</code> interface. */
class ColumnImpl : public Column {
    public:
        /** Creates a new instance of <code>ColumnImpl</code>. */
        ColumnImpl();

        // inherit doc
        virtual double getWidth() const;

        // inherit doc
        virtual void setWidth(double width);

    private:
        /** Stores the column width */
        double width;
};

}

#endif // COLUMNIMPL_H
