#pragma once
#include <windows.h>
#include <string>
#include <vector>

bool IsSupportedPath(const wchar_t* path);
HRESULT LaunchConversion(HMODULE module, const std::vector<std::wstring>& paths);

inline const wchar_t* MenuText(const wchar_t* traditionalChinese, const wchar_t* english) {
    return PRIMARYLANGID(GetUserDefaultUILanguage()) == LANG_CHINESE ? traditionalChinese : english;
}
