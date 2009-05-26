// File: IndexedCollectionImpl.h
// IndexedCollectionImpl declaration file
//

#ifndef INDEXEDCOLLECTIONIMPL_H
#define INDEXEDCOLLECTIONIMPL_H

#include "splib.h"
#include <map>
#include "splibint.h"

namespace splib {

/** Default implementation of the <code>IndexedCollection</code> interface. */
template<class T,class I>
class IndexedCollectionImpl : public IndexedCollection<T> {
    private:
        /**
         * An internal collection used to store and manage
         * the indexed collection of objects
         */
        typedef std::map<int, T*> Map;

        /** An iterator over the internal collection */
        typedef typename std::map<int, T*>::iterator MapIterator;

        /** A reverse iterator over the internal collection */
        typedef typename std::map<int, T*>::reverse_iterator MapReverseIterator;

        /** A shortcut to <code>splib::IndexedCollection<T>::Entry</code> */
        typedef typename splib::IndexedCollection<T>::Entry Entry;

        /** A shortcut to <code>splib::IndexedCollection<T>::Iterator</code> */
        typedef typename splib::IndexedCollection<T>::Iterator BaseIterator;

        /**
         * An implementation of
         * <code>splib::IndexedCollection<T>::Iterator</code>
         */
        class Iterator : public BaseIterator {
            public:
                /** Creates a new <code>Iterator</code> */
                Iterator(Map& m) : map(m) {
                    iterator = map.begin();
                }

                // inherit doc
                virtual bool hasNext() const {
                    return iterator != map.end();
                }

                // inherit doc
                virtual Entry next() {
                    if (!hasNext()) {
                        throw IllegalStateException(_T("no more objects"));
                    }
                    Entry entry(*iterator->second, iterator->first);
                    iterator++;
                    return entry;
                }

            private:
                /**
                 * Assignment operator. Declared private to disallow
                 * assignments.
                 */
                Iterator& operator = (const Iterator& i) {return *this;}

            private:
                /** A reference to the collection to iterate over */
                Map& map;

                /** The iterator used to implement the iteration */
                MapIterator iterator;
        };

    public:
        /** Creates a new <code>IndexedCollectionImpl</code> */
        IndexedCollectionImpl() {}

        /** Cleans up resources allocated by this object */
        virtual ~IndexedCollectionImpl() {
            for (MapIterator i = map.begin(); i != map.end(); i++) {
                delete i->second;
            }
        }

        // inherit doc
        virtual T& get(int index) {
            T*& obj = map[index];
            if (obj == 0) {
                obj = new I();
            }
            return *obj;
        }

        // inherit doc
        virtual bool contains(int index) {
            return map.find(index) != map.end();
        }

        // inherit doc
        virtual void remove(int index) {
            MapIterator i = map.find(index);
            if (i != map.end()) {
                delete i->second;
                map.erase(i);
            }
        }

        // inherit doc
        virtual int size() const {
            return (int)map.size();
        }

        // inherit doc
        virtual bool isEmpty() const {
            return map.empty();
        }

        // inherit doc
        virtual BaseIterator* iterator() {
            return new Iterator(map);
        }

        // inherit doc
        virtual Entry first() {
            if (size() == 0) {
                throw IllegalStateException(_T("collection is empty"));
            }
            MapIterator i = map.begin();
            return Entry(*i->second, i->first);
        }

        // inherit doc
        virtual Entry last() {
            if (size() == 0) {
                throw IllegalStateException(_T("collection is empty"));
            }
            MapReverseIterator i = map.rbegin();
            return Entry(*i->second, i->first);
        }

    private:
        /**
         * The object used to store and manage the indexed collection
         * of objects
         */
        Map map;
};

}

#endif // INDEXEDCOLLECTIONIMPL_H
