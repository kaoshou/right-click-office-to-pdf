#include "Launch.h"
#include <objbase.h>
#include <cstdint>
#include <memory>

bool IsSupportedPath(const wchar_t* path) {
    const auto extension = wcsrchr(path, L'.');
    if (!extension) return false;
    for (auto supported : {L".doc", L".docx", L".xls", L".xlsx", L".ppt",
                           L".pptx", L".odt", L".ods", L".odp"}) {
        if (!_wcsicmp(extension, supported)) return true;
    }
    return false;
}

namespace {
struct HandleCloser { void operator()(void* handle) const { CloseHandle(handle); } };
using Handle = std::unique_ptr<void, HandleCloser>;
struct AttributeListCloser {
    void operator()(LPPROC_THREAD_ATTRIBUTE_LIST list) const { DeleteProcThreadAttributeList(list); }
};
}

HRESULT LaunchConversion(HMODULE module, const std::vector<std::wstring>& paths) {
    if (paths.empty()) return E_INVALIDARG;
    wchar_t modulePath[32768], systemDir[MAX_PATH], tempDir[MAX_PATH + 1];
    const DWORD moduleLength = GetModuleFileNameW(module, modulePath, ARRAYSIZE(modulePath));
    if (!moduleLength || moduleLength >= ARRAYSIZE(modulePath)) return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    const DWORD systemLength = GetSystemDirectoryW(systemDir, ARRAYSIZE(systemDir));
    const DWORD tempLength = GetTempPathW(ARRAYSIZE(tempDir), tempDir);
    if (!systemLength || systemLength >= ARRAYSIZE(systemDir) || !tempLength || tempLength >= ARRAYSIZE(tempDir))
        return HRESULT_FROM_WIN32(ERROR_BAD_PATHNAME);
    const std::wstring directory(modulePath, wcsrchr(modulePath, L'\\') - modulePath);
    const std::wstring script = directory + L"\\Invoke-Selection.ps1";
    const std::wstring powershell = std::wstring(systemDir) + L"\\WindowsPowerShell\\v1.0\\powershell.exe";

    GUID id{};
    HRESULT hr = CoCreateGuid(&id);
    if (FAILED(hr)) return hr;
    wchar_t idText[40];
    StringFromGUID2(id, idText, ARRAYSIZE(idText));
    const auto temp = std::wstring(tempDir) + L"RightClickOfficeToPDF-" + idText + L".tmp";
    SECURITY_ATTRIBUTES security{sizeof(security), nullptr, TRUE};
    HANDLE raw = CreateFileW(temp.c_str(), GENERIC_READ | GENERIC_WRITE, 0, &security, CREATE_NEW,
                             FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, nullptr);
    if (raw == INVALID_HANDLE_VALUE) return HRESULT_FROM_WIN32(GetLastError());
    Handle selection(raw);

    // Pass the selection as data. The kernel removes it when the last handle
    // closes, including when either process fails or PowerShell cannot start.
    std::wstring data(1, L'\xFEFF');
    for (const auto& path : paths) {
        if (path.find_first_of(L"\r\n") != std::wstring::npos) return E_INVALIDARG;
        data.append(path).append(L"\r\n");
    }
    if (data.size() > MAXDWORD / sizeof(wchar_t)) return E_OUTOFMEMORY;
    DWORD written = 0, bytes = static_cast<DWORD>(data.size() * sizeof(wchar_t));
    if (!WriteFile(raw, data.data(), bytes, &written, nullptr)) return HRESULT_FROM_WIN32(GetLastError());
    if (written != bytes) return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
    if (!SetFilePointerEx(raw, {}, nullptr, FILE_BEGIN)) return HRESULT_FROM_WIN32(GetLastError());

    // Only this handle is inherited from the COM surrogate.
    SIZE_T size = 0;
    InitializeProcThreadAttributeList(nullptr, 1, 0, &size);
    std::vector<BYTE> storage(size);
    auto attributes = reinterpret_cast<LPPROC_THREAD_ATTRIBUTE_LIST>(storage.data());
    if (!InitializeProcThreadAttributeList(attributes, 1, 0, &size)) return HRESULT_FROM_WIN32(GetLastError());
    std::unique_ptr<std::remove_pointer_t<LPPROC_THREAD_ATTRIBUTE_LIST>, AttributeListCloser> cleanup(attributes);
    if (!UpdateProcThreadAttribute(attributes, 0, PROC_THREAD_ATTRIBUTE_HANDLE_LIST,
                                   &raw, sizeof(raw), nullptr, nullptr)) return HRESULT_FROM_WIN32(GetLastError());
    STARTUPINFOEXW startup{};
    startup.StartupInfo.cb = sizeof(startup);
    startup.lpAttributeList = attributes;
    PROCESS_INFORMATION process{};
    std::wstring args = L"\"" + powershell + L"\" -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File \"" +
                        script + L"\" " + std::to_wstring(reinterpret_cast<uintptr_t>(raw));
    if (!CreateProcessW(powershell.c_str(), args.data(), nullptr, nullptr, TRUE,
                        CREATE_NO_WINDOW | EXTENDED_STARTUPINFO_PRESENT, nullptr, directory.c_str(),
                        &startup.StartupInfo, &process)) return HRESULT_FROM_WIN32(GetLastError());
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    return S_OK;
}
