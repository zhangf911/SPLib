// File: ByteArray.h
// ByteArray declaration file
//

#ifndef BYTEARRAY_H
#define BYTEARRAY_H

namespace splib {

/** An array of bytes that grows automatically as necessary. */
class ByteArray {
    public:
        /** Creates a new <code>ByteArray</code>. */
        ByteArray();

        /** Cleans up resources allocated by this object. */
        virtual ~ByteArray();

        /**
         * Adds a byte to the end of this array; grows the array
         * if necessary.
         * @param b byte to add
         * @return a reference to this object
         */
        ByteArray& write(unsigned char b);

        /**
         * Adds a number of bytes from a C-style array
         * to the end of this array.
         * @param b a pointer to the array of bytes to add
         * @param count the number of bytes to add
         * @return a reference to this object
         */
        ByteArray& write(const unsigned char* p, int count);
        
        /**
         * Returns a pointer to the location where all bytes kept
         * in this byte array are stored.
         * @return a pointer to the location where all bytes kept
         * in this byte array are stored
         */
        unsigned char* data() const;
        
        /**
         * Returnes the number of bytes in this byte array.
         * @return the number of bytes in this byte array
         */
        int size() const;

    private:
        /** Extends the internal buffer to the given number of bytes. */
        void extend(int s);

    private:
        /** The buffer that stores bytes */
        unsigned char* buffer;

        /** The number of bytes currently stored in this byte array */
        int arraySize;

        /** The size of the buffer */
        int bufferSize;
};

}

#endif // BYTEARRAY_H
