// File: ByteArray.cpp
// ByteArray implementation file
//

#include "ByteArray.h"
#include <memory.h>
#include "splibint.h"

namespace splib {

ByteArray::ByteArray() {
    arraySize = 0;
    bufferSize = 16;
    buffer = new unsigned char[bufferSize];
}

ByteArray::~ByteArray() {
    delete [] buffer;
}

ByteArray& ByteArray::write(unsigned char b) {
    if (arraySize + 1 > bufferSize) {
        extend(arraySize + 1);
    }
    buffer[arraySize++] = b;
    return *this;
}

ByteArray& ByteArray::write(const unsigned char* b, int count) {
    if (arraySize + count > bufferSize) {
        extend(arraySize + count);
    }
    ::memcpy(buffer + arraySize, b, count);
    arraySize += count;
    return *this;
}

unsigned char* ByteArray::data() const {
    return buffer;
}

int ByteArray::size() const {
    return arraySize;
}

void ByteArray::extend(int s) {
    bufferSize = s * 2;
    unsigned char* tmp = new unsigned char[bufferSize];
    memcpy(tmp, buffer, arraySize);
    delete [] buffer;
    buffer = tmp;
}

}
