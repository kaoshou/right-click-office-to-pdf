#include "Launch.h"
#include <shobjidl.h>
#include <shlwapi.h>
#include <wrl.h>
#include <atomic>
#include <memory>

using namespace Microsoft::WRL;
extern "C" IMAGE_DOS_HEADER __ImageBase;
namespace {
constexpr GUID CommandId{0x91d6a591, 0x2bd6, 0x4452, {0xb6, 0x6b, 0x99, 0x7e, 0x66, 0xb0, 0x8c, 0x4b}};
std::atomic<long> active{0};
struct ObjectCount {
    ObjectCount() { ++active; }
    ~ObjectCount() { --active; }
};

// Read Shell metadata only. Do not open or hydrate the selected documents.
HRESULT GetPaths(IShellItemArray* items, std::vector<std::wstring>* paths) {
    if (!items) return S_FALSE;
    DWORD count = 0;
    HRESULT hr = items->GetCount(&count);
    if (FAILED(hr) || !count) return FAILED(hr) ? hr : S_FALSE;
    for (DWORD i = 0; i < count; ++i) {
        ComPtr<IShellItem> item;
        hr = items->GetItemAt(i, &item);
        if (FAILED(hr)) return hr;
        PWSTR raw = nullptr;
        hr = item->GetDisplayName(SIGDN_FILESYSPATH, &raw);
        if (FAILED(hr)) return hr;
        std::unique_ptr<wchar_t, decltype(&CoTaskMemFree)> path(raw, CoTaskMemFree);
        if (!IsSupportedPath(raw)) return S_FALSE;
        SFGAOF attributes = 0;
        hr = item->GetAttributes(SFGAO_FOLDER | SFGAO_FILESYSTEM, &attributes);
        if (FAILED(hr)) return hr;
        if (!(attributes & SFGAO_FILESYSTEM) || (attributes & SFGAO_FOLDER)) return S_FALSE;
        if (paths) paths->emplace_back(raw);
    }
    return S_OK;
}

class Command final : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand>, private ObjectCount {
public:
    IFACEMETHODIMP GetTitle(IShellItemArray*, LPWSTR* title) override {
        return SHStrDupW(L"轉換為 PDF 檔案", title);
    }
    IFACEMETHODIMP GetIcon(IShellItemArray*, LPWSTR* icon) override {
        return SHStrDupW(L"shell32.dll,70", icon);
    }
    IFACEMETHODIMP GetToolTip(IShellItemArray*, LPWSTR* tip) override { *tip = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetCanonicalName(GUID* name) override { *name = CommandId; return S_OK; }
    IFACEMETHODIMP GetState(IShellItemArray* items, BOOL okToBeSlow, EXPCMDSTATE* state) override {
        *state = ECS_HIDDEN;
        if (!items) return S_OK;
        DWORD count = 0;
        if (!okToBeSlow && SUCCEEDED(items->GetCount(&count)) && count > 16) return E_PENDING;
        if (GetPaths(items, nullptr) == S_OK) *state = ECS_ENABLED;
        return S_OK;
    }
    IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) override {
        try {
            std::vector<std::wstring> paths;
            HRESULT hr = GetPaths(items, &paths);
            if (hr != S_OK) return FAILED(hr) ? hr : E_INVALIDARG;
            return LaunchConversion(reinterpret_cast<HMODULE>(&__ImageBase), paths);
        } catch (const std::bad_alloc&) { return E_OUTOFMEMORY; }
        catch (...) { return E_FAIL; }
    }
    IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) override { *flags = ECF_DEFAULT; return S_OK; }
    IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** commands) override { *commands = nullptr; return E_NOTIMPL; }
};

class Factory final : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IClassFactory>, private ObjectCount {
public:
    IFACEMETHODIMP CreateInstance(IUnknown* outer, REFIID iid, void** object) override {
        *object = nullptr;
        if (outer) return CLASS_E_NOAGGREGATION;
        auto command = Make<Command>();
        return command ? command.CopyTo(iid, object) : E_OUTOFMEMORY;
    }
    IFACEMETHODIMP LockServer(BOOL lock) override {
        if (lock) ++active; else --active;
        return S_OK;
    }
};
}

STDAPI DllGetClassObject(REFCLSID clsid, REFIID iid, void** object) {
    *object = nullptr;
    if (clsid != CommandId) return CLASS_E_CLASSNOTAVAILABLE;
    auto factory = Make<Factory>();
    return factory ? factory.CopyTo(iid, object) : E_OUTOFMEMORY;
}
STDAPI DllCanUnloadNow() { return active == 0 ? S_OK : S_FALSE; }
