#include "Launch.h"
#include <shellapi.h>

// The package executable accepts file paths as well as a normal app activation.
int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int) {
    try {
        int argc = 0;
        auto argv = CommandLineToArgvW(GetCommandLineW(), &argc);
        if (!argv) return 1;
        std::vector<std::wstring> paths;
        for (int i = 1; i < argc; ++i) {
            if (IsSupportedPath(argv[i])) paths.emplace_back(argv[i]);
        }
        LocalFree(argv);
        if (!paths.empty()) {
            HRESULT hr = LaunchConversion(instance, paths);
            if (FAILED(hr)) {
                MessageBoxW(nullptr, MenuText(L"無法啟動轉檔工具，請確認安裝完整後再試一次。", L"Could not start the converter. Please check the installation."),
                            L"Right Click Office to PDF", MB_OK | MB_ICONERROR);
            }
            return FAILED(hr);
        }
        MessageBoxW(nullptr, MenuText(L"請在檔案總管選取文件，按右鍵，再選擇「轉換為 PDF 檔案」。", L"Right-click a document in File Explorer and choose Convert to PDF."),
                    L"Right Click Office to PDF", MB_OK | MB_ICONINFORMATION);
        return 0;
    } catch (...) { return 1; }
}
