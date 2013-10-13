// File: splibint.h
// Internal splib declarations
//

#ifndef SPLIBINT_H
#define SPLIBINT_H

#ifndef WIN32
#define _T(x) x
#define _stprintf sprintf
#define _stscanf  sscanf
#define _istdigit isdigit
#define _tcslen   strlen
#define _tcscmp   strcmp
#define _tcsncmp  strncmp
#define _tfopen   fopen
#endif

// platform-specific includes and declarations
#ifdef WIN32
    #include <tchar.h>
// #else
//    typedef char _TCHAR;
#endif // WIN32

#if defined(_DEBUG) && defined(WIN32)
#include <crtdbg.h>
#define new new(_NORMAL_BLOCK, __FILE__, __LINE__)
#endif // _DEBUG && WIN32

#endif // SPLIBINT_H
