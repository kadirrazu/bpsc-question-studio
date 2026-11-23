/* Run embedded Python inside the branded Windows application process. */
#ifndef UNICODE
#define UNICODE
#endif
#define _UNICODE
#include <windows.h>
#include <wchar.h>

typedef int (*PythonMain)(int, wchar_t **);

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE previous, PWSTR arguments, int show) {
    wchar_t executable[32768], directory[32768], runtime[32768], dll[32768], script[32768];
    DWORD length = GetModuleFileNameW(NULL, executable, ARRAYSIZE(executable));
    if (!length || length >= ARRAYSIZE(executable)) return 1;
    wcscpy(directory, executable);
    wchar_t *separator = wcsrchr(directory, L'\\');
    if (!separator) return 1;
    *separator = 0;
    if (swprintf(runtime, ARRAYSIZE(runtime), L"%ls\\Runtime", directory) < 0 ||
        swprintf(dll, ARRAYSIZE(dll), L"%ls\\python312.dll", runtime) < 0 ||
        swprintf(script, ARRAYSIZE(script), L"%ls\\main.py", directory) < 0) return 1;

    /* Python extensions may also import the stable ABI shim python3.dll. */
    DLL_DIRECTORY_COOKIE runtime_search = AddDllDirectory(runtime);
    HMODULE python = runtime_search ?
        LoadLibraryExW(dll, NULL, LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_DEFAULT_DIRS) : NULL;
    PythonMain run = python ? (PythonMain)GetProcAddress(python, "Py_Main") : NULL;
    if (!run || !SetCurrentDirectoryW(directory)) {
        MessageBoxW(NULL, L"Extract the entire ZIP before running BPSC Question Studio.",
                    L"BPSC Question Studio", MB_OK | MB_ICONERROR);
        return 1;
    }
    wchar_t *argv[] = {executable, script, NULL};
    return run(2, argv);
}
