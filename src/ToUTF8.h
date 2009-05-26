// File: ToUTF8.h
// ToUTF8 declaration file
//

#ifndef TOUTF8_H
#define TOUTF8_H

#include "splib.h"

namespace splib {

/**
 * A convenience utility for converting a <code>_TCHAR</code> string to
 * a UTF8-encoded <code>char</code> string. This object receives a
 * <code>_TCHAR</code> string in its constructor. Once the object is
 * constructed, the resulting UTF8-encoded <code>char</code> string
 * can be retrieved using the <code>get()</code> method.
 * <p>
 * <code>ToUTF8</code> behaves differently on different platforms:
 * <ul>
 * <li>On Windows, if _UNICODE is defined, the input string is converted
 *     to UTF8 by calling <code>WideCharToMultiByte()</code> with the code
 *     page parameter set to CP_UTF8.
 * <li>On Windows, if _UNICODE is not defined, the input string is first
 *     converted to UTF16 by calling <code>MultiByteToWideChar()</code>
 *     with the code page parameter set to CP_ACP, and then to UTF8 by
 *     calling <code>WideCharToMultiByte()</code> with the code
 *     page parameter set to CP_UTF8.
 * <li>On other platforms, the input string is assumed to be UTF8 already,
 *     therefore no conversion is performed.
 * </ul>
 * <p>
 * Make sure you use the string obtained from this object before this
 * object goes away and before the original string passed to this object
 * goes away.
 */
class ToUTF8 {
    public:
        /** Creates a <code>ToUTF8</code> object. */
        ToUTF8(const _TCHAR* s);

        /** Destructor */
        virtual ~ToUTF8();

        /** Retrieves the UTF8-encoded string. */
        const char* get() const;

        /**
         * Retrieves the number of characters in the resulting string,
         * not including terminating null character.
         */
        int length() const;

    private:
        /** The string stored in this object */
        const char* buffer;

        /**
         * Indicates whether this object uses its own buffer or the
         * original string
         */
        bool ownBuffer;
};

}

#endif // TOUTF8_H
