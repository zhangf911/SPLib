// File: RowsImpl.h
// RowsImpl declaration file
//

#ifndef ROWSIMPL_H
#define ROWSIMPL_H

#include "splib.h"
#include "IndexedCollectionImpl.h"
#include "RowImpl.h"

namespace splib {

/** Default implementation of the <code>Rows</code> interface. */
class RowsImpl : public Rows {
    public:
        // inherit doc
        virtual Row& get(int index);

        // inherit doc
        virtual bool contains(int index);

        // inherit doc
        virtual void remove(int index);

        // inherit doc
        virtual int size() const;

        // inherit doc
        virtual bool isEmpty() const;

        // inherit doc
        virtual Iterator* iterator();

        // inherit doc
        virtual Entry first();

        // inherit doc
        virtual Entry last();

    private:
        /** The object used to store and manage the collection of rows */
        IndexedCollectionImpl<Row,RowImpl> collection;
};

}

#endif // ROWSIMPL_H
