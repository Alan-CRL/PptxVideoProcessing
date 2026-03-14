module;

#define NOMINMAX
#include <Windows.h>
#include <conio.h>
#include <cstdio>
#include <fcntl.h>
#include <io.h>

module PptxVideoProcessing.Helper.Console;

import std.compat;

namespace
{
    constexpr std::size_t kProgressSafetyPadding = 6;

    std::size_t g_last_progress_width = 0;
    bool g_progress_active = false;
    bool g_has_progress_origin = false;
    COORD g_progress_origin{};

    [[nodiscard]] HANDLE GetConsoleOutputHandle() noexcept
    {
        return ::GetStdHandle(STD_OUTPUT_HANDLE);
    }

    [[nodiscard]] bool TryGetConsoleInfo(CONSOLE_SCREEN_BUFFER_INFO& info) noexcept
    {
        const HANDLE handle = GetConsoleOutputHandle();
        return handle != nullptr && handle != INVALID_HANDLE_VALUE && ::GetConsoleScreenBufferInfo(handle, &info) != FALSE;
    }

    [[nodiscard]] bool IsCombiningCharacter(wchar_t character) noexcept
    {
        return (character >= 0x0300 && character <= 0x036F) || (character >= 0x1AB0 && character <= 0x1AFF) ||
               (character >= 0x1DC0 && character <= 0x1DFF) || (character >= 0x20D0 && character <= 0x20FF) ||
               (character >= 0xFE20 && character <= 0xFE2F);
    }

    [[nodiscard]] bool IsWideCharacter(wchar_t character) noexcept
    {
        return (character >= 0x1100 && character <= 0x115F) || (character >= 0x2E80 && character <= 0xA4CF) ||
               (character >= 0xAC00 && character <= 0xD7A3) || (character >= 0xF900 && character <= 0xFAFF) ||
               (character >= 0xFE10 && character <= 0xFE19) || (character >= 0xFE30 && character <= 0xFE6F) ||
               (character >= 0xFF01 && character <= 0xFF60) || (character >= 0xFFE0 && character <= 0xFFE6);
    }

    [[nodiscard]] std::size_t ConsoleCellWidth(wchar_t character) noexcept
    {
        if (character == L'\0' || character == L'\r' || character == L'\n')
        {
            return 0;
        }

        if (character < 0x20 || IsCombiningCharacter(character))
        {
            return 0;
        }

        return IsWideCharacter(character) ? 2U : 1U;
    }

    [[nodiscard]] std::size_t MeasureConsoleDisplayWidth(std::wstring_view text) noexcept
    {
        std::size_t width = 0;

        for (const wchar_t character : text)
        {
            width += ConsoleCellWidth(character);
        }

        return width;
    }

    [[nodiscard]] std::wstring TruncateProgressMessage(std::wstring_view message, std::size_t max_columns)
    {
        constexpr std::wstring_view ellipsis = L"...";
        const std::size_t ellipsis_width = MeasureConsoleDisplayWidth(ellipsis);

        if (max_columns == 0)
        {
            return {};
        }

        if (MeasureConsoleDisplayWidth(message) <= max_columns)
        {
            return std::wstring(message);
        }

        if (max_columns <= ellipsis_width)
        {
            return std::wstring(ellipsis.substr(0, std::min<std::size_t>(ellipsis.size(), max_columns)));
        }

        const std::size_t target_width = max_columns - ellipsis_width;
        std::wstring truncated;
        std::size_t current_width = 0;

        for (const wchar_t character : message)
        {
            const std::size_t character_width = ConsoleCellWidth(character);

            if (current_width + character_width > target_width)
            {
                break;
            }

            truncated.push_back(character);
            current_width += character_width;
        }

        truncated.append(ellipsis);
        return truncated;
    }

    [[nodiscard]] std::size_t GetAvailableProgressColumns(const CONSOLE_SCREEN_BUFFER_INFO& info) noexcept
    {
        const SHORT window_width = static_cast<SHORT>(info.srWindow.Right - info.srWindow.Left + 1);
        const SHORT origin_x = g_has_progress_origin ? g_progress_origin.X : info.dwCursorPosition.X;

        if (window_width <= 0 || origin_x < 0 || origin_x >= window_width)
        {
            return 0;
        }

        std::size_t available = static_cast<std::size_t>(window_width - origin_x);

        if (available > kProgressSafetyPadding)
        {
            available -= kProgressSafetyPadding;
        }
        else if (available > 1)
        {
            --available;
        }

        return available;
    }

    [[nodiscard]] std::wstring FitProgressMessage(std::wstring_view message)
    {
        CONSOLE_SCREEN_BUFFER_INFO info{};

        if (!TryGetConsoleInfo(info))
        {
            return std::wstring(message);
        }

        const std::size_t max_columns = GetAvailableProgressColumns(info);

        if (max_columns == 0)
        {
            return std::wstring(message);
        }

        return TruncateProgressMessage(message, max_columns);
    }

    bool TryWriteProgressLine(std::wstring_view message)
    {
        CONSOLE_SCREEN_BUFFER_INFO info{};

        if (!TryGetConsoleInfo(info))
        {
            return false;
        }

        const HANDLE handle = GetConsoleOutputHandle();
        const std::wstring fitted_message = FitProgressMessage(message);

        if (!g_has_progress_origin)
        {
            g_progress_origin = info.dwCursorPosition;
            g_has_progress_origin = true;
        }

        ::SetConsoleCursorPosition(handle, g_progress_origin);

        DWORD written = 0;
        if (!fitted_message.empty())
        {
            ::WriteConsoleW(handle, fitted_message.data(), static_cast<DWORD>(fitted_message.size()), &written, nullptr);
        }

        const std::size_t fitted_width = MeasureConsoleDisplayWidth(fitted_message);

        if (fitted_width < g_last_progress_width)
        {
            const std::wstring padding(g_last_progress_width - fitted_width, L' ');
            ::WriteConsoleW(handle, padding.data(), static_cast<DWORD>(padding.size()), &written, nullptr);
        }

        g_last_progress_width = fitted_width;
        g_progress_active = true;
        return true;
    }

    void FinishProgressLineImpl()
    {
        if (!g_progress_active)
        {
            return;
        }

        DWORD written = 0;
        const HANDLE handle = GetConsoleOutputHandle();

        if (handle != nullptr && handle != INVALID_HANDLE_VALUE)
        {
            ::WriteConsoleW(handle, L"\r\n", 2, &written, nullptr);
        }
        else
        {
            std::wcout << L'\n';
            std::wcout.flush();
        }

        g_last_progress_width = 0;
        g_progress_active = false;
        g_has_progress_origin = false;
    }
}

namespace pptxvp::helper
{
    void InitializeUtf16Console()
    {
        static bool initialized = false;

        if (initialized)
        {
            return;
        }

        _setmode(_fileno(stdin), _O_U16TEXT);
        _setmode(_fileno(stdout), _O_U16TEXT);
        _setmode(_fileno(stderr), _O_U16TEXT);

        initialized = true;
    }

    void WriteLine(std::wstring_view message)
    {
        FinishProgressLineImpl();
        std::wcout << message << L'\n';
    }

    void WriteErrorLine(std::wstring_view message)
    {
        FinishProgressLineImpl();
        std::wcerr << message << L'\n';
    }

    void WriteProgressLine(std::wstring_view message)
    {
        std::wcout.flush();
        std::wcerr.flush();

        if (TryWriteProgressLine(message))
        {
            return;
        }

        const std::wstring fitted_message = FitProgressMessage(message);
        std::wcout << L'\r' << fitted_message;

        const std::size_t fitted_width = MeasureConsoleDisplayWidth(fitted_message);

        if (fitted_width < g_last_progress_width)
        {
            std::wcout << std::wstring(g_last_progress_width - fitted_width, L' ');
        }

        std::wcout.flush();
        g_last_progress_width = fitted_width;
        g_progress_active = true;
    }

    void FinishProgressLine()
    {
        FinishProgressLineImpl();
    }

    void WaitForAnyKey(std::wstring_view prompt)
    {
        FinishProgressLineImpl();

        if (!prompt.empty())
        {
            std::wcout << prompt;
            std::wcout.flush();
        }

        (void)_getwch();
        std::wcout << L'\n';
    }
}
