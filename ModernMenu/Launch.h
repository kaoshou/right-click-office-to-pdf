#pragma once
#include <windows.h>
#include <string>
#include <vector>

bool IsSupportedPath(const wchar_t* path);
HRESULT LaunchConversion(HMODULE module, const std::vector<std::wstring>& paths);
