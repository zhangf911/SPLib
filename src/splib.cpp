// File: splib.cpp
// General splib code
//

#if defined(_DEBUG) && defined(WIN32)

#include <windows.h>
#include "splibint.h"

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call,
                      LPVOID lpReserved) {
    hModule, lpReserved;
	switch (ul_reason_for_call)	{
        case DLL_PROCESS_ATTACH:
            //_CrtSetDbgFlag(_CrtSetDbgFlag(_CRTDBG_REPORT_FLAG)
            //    | _CRTDBG_LEAK_CHECK_DF | _CRTDBG_ALLOC_MEM_DF);
            break;
        case DLL_THREAD_ATTACH:
            break;
        case DLL_THREAD_DETACH:
            break;
        case DLL_PROCESS_DETACH:
            _CrtDumpMemoryLeaks();
            break;
    }
    return TRUE;
}

#endif // _DEBUG && WIN32
