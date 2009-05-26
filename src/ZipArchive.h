// File: ZipArchive.h
// ZipArchive declaration file
//

#ifndef ZIPARCHIVE_H
#define ZIPARCHIVE_H

#include "splib.h"
#include "zip.h"

namespace splib {

/** A writable zip archive. */
class ZipArchive {
    public:
        /**
         * Creates a new zip archive and opens it for writing. If the file
         * with the specified name already exists, its contents is cleared.
         */
        ZipArchive(const _TCHAR* pathname);

        /**
         * Destructor. Closes the currently opened entry if it is not closed
         * yet. Closes the archive if it is not closed yet.
         */
        virtual ~ZipArchive();

        /**
         * Opens an entry in this zip archive and makes it ready for writing
         * via the <code>write()</code> method.
         */
        void openEntry(const char* entryName);

        /** Writes data into the currently opened entry. */
        void write(const void* buffer, unsigned length);

        /**
         * Finishes writing into the currently opened entry and closes
         * the entry.
         */
        void closeEntry();

        /**
         * Closes this zip archive. Closes the currently opened entry, if
         * there is one.
         */
        void close();

    private:
        /** A handle to the underlying zip archive */
        zipFile zf;

        /** Indicates whether there is a currently opened entry */
        bool entryOpened;
};

// convenience operators
ZipArchive& operator << (ZipArchive& ar, int val);
ZipArchive& operator << (ZipArchive& ar, long val);
ZipArchive& operator << (ZipArchive& ar, double val);
ZipArchive& operator << (ZipArchive& ar, const char* val);

}

#endif // ZIPARCHIVE_H
