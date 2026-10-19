// Прикладная библиотека КОМПАС-3D «Проверка чертежа» (нативная, без регистрации COM).
// Пункты меню запускают drawcheck.py --menu N; вся логика — в Python.
#include <windows.h>
#include <string>

#define IDR_LIBID 2000

static HMODULE g_module = nullptr;

BOOL APIENTRY DllMain(HMODULE h, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH) g_module = h;
    return TRUE;
}

static std::wstring Here()
{
    wchar_t buf[MAX_PATH];
    GetModuleFileNameW(g_module, buf, MAX_PATH);
    std::wstring p(buf);
    return p.substr(0, p.find_last_of(L"\\/"));
}

static bool Exists(const std::wstring& p)
{
    return GetFileAttributesW(p.c_str()) != INVALID_FILE_ATTRIBUTES;
}

// путь к python: общая папка плагинов ..\Python (ставится один раз на все плагины),
// иначе python.txt рядом с библиотекой, иначе python из PATH
static std::wstring PythonExe(bool console)
{
    std::wstring exe = console ? L"python.exe" : L"pythonw.exe";
    std::wstring shared = Here() + L"\\..\\Python\\" + exe;
    if (Exists(shared)) return shared;
    HANDLE f = CreateFileW((Here() + L"\\python.txt").c_str(), GENERIC_READ, FILE_SHARE_READ,
                           nullptr, OPEN_EXISTING, 0, nullptr);
    if (f != INVALID_HANDLE_VALUE) {
        char buf[1024] = {};
        DWORD n = 0;
        ReadFile(f, buf, sizeof(buf) - 1, &n, nullptr);
        CloseHandle(f);
        wchar_t w[1024] = {};
        MultiByteToWideChar(CP_UTF8, 0, buf, (int)n, w, 1023);
        std::wstring p(w);
        p.erase(p.find_last_not_of(L" \r\n\t") + 1);
        if (!p.empty()) {
            std::wstring dir = p.substr(0, p.find_last_of(L"\\/"));
            std::wstring cand = dir + (console ? L"\\python.exe" : L"\\pythonw.exe");
            exe = Exists(cand) ? cand : p;
        }
    }
    return exe;
}

static std::wstring Script()
{
    std::wstring s = Here() + L"\\drawcheck.py";
    if (!Exists(s)) s = Here() + L"\\..\\drawcheck.py";
    return s;
}

// Регистрирует объект КОМПАС, в котором работает библиотека, в ROT под уникальным
// именем — чтобы скрипт подключился именно к этому окну, а не к первому запущенному КОМПАС.
static std::wstring PublishHostApp()
{
    static DWORD cookie = 0;
    static IUnknown* published = nullptr;
    std::wstring name = L"DrawCheck.Kompas." + std::to_wstring(GetCurrentProcessId());

    typedef IDispatch* (WINAPI *CreateApp)();
    HMODULE k7 = GetModuleHandleW(L"kAPI7.dll");
    if (!k7) k7 = LoadLibraryW(L"kAPI7.dll");
    CreateApp create = k7 ? (CreateApp)GetProcAddress(k7, "CreateKompasApplication") : nullptr;
    if (!create) return L"";

    IRunningObjectTable* rot = nullptr;
    if (FAILED(GetRunningObjectTable(0, &rot))) return L"";
    if (cookie) {               // повторный вызов — перерегистрировать
        rot->Revoke(cookie);
        cookie = 0;
        if (published) { published->Release(); published = nullptr; }
    }
    IDispatch* app = create();
    IMoniker* mk = nullptr;
    if (app && SUCCEEDED(CreateItemMoniker(L"!", name.c_str(), &mk))) {
        if (SUCCEEDED(rot->Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, app, mk, &cookie)))
            published = app;
        else
            app->Release();
        mk->Release();
    }
    rot->Release();
    return cookie ? name : L"";
}

extern "C" {

unsigned int WINAPI LIBRARYID() { return IDR_LIBID; }

LPWSTR WINAPI LIBRARYNAMEW()
{
    static wchar_t name[] = L"Проверка чертежа";
    return name;
}

void WINAPI LIBRARYENTRY(unsigned int comm)
{
    bool console = (comm == 2);  // папка чертежей — с окном хода работы
    std::wstring cmd = L"\"" + PythonExe(console) + L"\" \"" + Script() + L"\" --menu " + std::to_wstring(comm);
    std::wstring rotName = PublishHostApp();
    if (!rotName.empty()) cmd += L" --rot " + rotName;

    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = {};
    SetEnvironmentVariableW(L"PYTHONIOENCODING", L"utf-8");
    std::wstring dir = Here();
    if (!CreateProcessW(nullptr, &cmd[0], nullptr, nullptr, FALSE,
                        console ? CREATE_NEW_CONSOLE : CREATE_NO_WINDOW,
                        nullptr, dir.c_str(), &si, &pi)) {
        std::wstring msg = L"Не удалось запустить Python:\n" + cmd;
        MessageBoxW(nullptr, msg.c_str(), L"Проверка чертежа", MB_ICONERROR);
        return;
    }
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
}

}
