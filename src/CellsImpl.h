// File: CellsImpl.h
// CellsImpl declaration file
//

#ifndef CELLSIMPL_H
#define CELLSIMPL_H

#include "splib.h"
#include "IndexedCollectionImpl.h"
#include "CellImpl.h"

namespace splib {

/** Default implementation of the <code>Cells</code> interface. */
class CellsImpl : public Cells {
    public:
        // inherit doc
        virtual Cell& get(int index);

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
        /** The object used to manage the collection of cells */
        IndexedCollectionImpl<Cell,CellImpl> collection;
};

}

#endif // CELLSIMPL_H
