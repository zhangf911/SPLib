// File: ToUTF16.h
// ToUTF16 declaration file
//

#ifndef TOUTF16_H
#define TOUTF16_H

#include "splib.h"

namespace splib {

/**
 * A convenience utility for converting a <code>_TCHAR</code> string to
 * a UTF16-encoded string. This object receives a
 * <code>_TCHAR</code> string in its constructor. Once the object is
 * constructed, the resulting UTF16-encoded string can be retrieved
 * using the <code>get()</code> method.
 * <p>
 * Note that the resulting string type is <code>unsigned char*</code>.
 * This is because <code>wchar_t</code> is not usable here since it is
 * 16 bits on Windows and (usually) 32 bits on Linux.
 * <p>
 * <code>ToUTF16</code> behaves differently on different platforms:
 * <ul>
 * <li>On Windows, if _UNICODE is defined, the input string is UTF16
 *     already, therefore no conversion is performed.
 * <li>On Windows, if _UNICODE is not defined, the input string is
 *     converted to UTF16 by calling <code>MultiByteToWideChar()</code>
 *     with the code page parameter set to CP_ACP.
 * <li>On other platforms, the input string is converted to UTF16
 *     using the iconv package with the converter created by calling
 *     <code>iconv_open("UTF-16", "")</code>.
 * </ul>
 * <p>
 * Make sure you use the string obtained from this object before this
 * object goes away and before the original string passed to this object
 * goes away.
 */
class ToUTF16 {
    public:
        /** Creates a <code>ToUTF16</code> object. */
        ToUTF16(const _TCHAR* s);

        /** Destructor. */
        virtual ~ToUTF16();

        /** Retrieves the UTF16-encoded string. */
        const unsigned char* get() const;

        /**
         * Retrieves the number of 16-bit characters in the UTF16-encoded
         * string, not including terminating null character.
         */
        int length() const;

    private:
        /** A pointer to the string stored in this object. */
        const unsigned char* buffer;

        /**
         * The number of UTF16 characters in the string pointed
         * by <code>buffer</code>
         */
        int len;

        /**
         * Indicates whether this object uses its own buffer or the
         * original string
         */
        bool ownBuffer;
};

}

#endif // TOUTF16_H
