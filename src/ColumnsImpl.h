// File: ColumnsImpl.h
// ColumnsImpl declaration file
//

#ifndef COLUMNSIMPL_H
#define COLUMNSIMPL_H

#include "splib.h"
#include "IndexedCollectionImpl.h"
#include "ColumnImpl.h"

namespace splib {

/** Default implementation of the <code>Columns</code> interface. */
class ColumnsImpl : public Columns {
    public:
        // inherit doc
        virtual Column& get(int index);

        // inherit doc
        virtual Column& get(const _TCHAR* column);

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
        /** The object used to manage the collection of columns */
        IndexedCollectionImpl<Column,ColumnImpl> collection;
};

}

#endif // COLUMNSIMPL_H
