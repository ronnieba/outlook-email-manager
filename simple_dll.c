
#include <windows.h>

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
    switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
        // ����� ���� ���
        HANDLE hFile = CreateFileA("C:\\Users\\ronni\\AppData\\Local\\Temp\\simple_dll_loaded.txt",
                                   GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE) {
            const char* message = "Simple DLL Loaded Successfully!";
            DWORD bytesWritten;
            WriteFile(hFile, message, strlen(message), &bytesWritten, NULL);
            CloseHandle(hFile);
        }
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}
