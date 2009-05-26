// File: RowImpl.h
// RowImpl declaration file
//

#ifndef ROWIMPL_H
#define ROWIMPL_H

#include "splib.h"
#include "CellsImpl.h"

namespace splib {

/** Default implementation of the <code>Row</code> interface. */
class RowImpl : public Row {
    public:
        /** Creates a new <code>RowImpl</code> */
        RowImpl();

        // inherit doc
        virtual double getHeight() const;

        // inherit doc
        virtual void setHeight(double height);

        // inherit doc
        virtual Cells& cells();

    private:
        /** Stores the height */
        double height;

        /** The collection of cells */
        CellsImpl cellsImpl;
};

}

#endif // ROWIMPL_H
