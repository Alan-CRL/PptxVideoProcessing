module;

#define NOMINMAX
#include <Windows.h>

module PptxVideoProcessing.Helper.Utf;

import std.compat;

namespace
{
    [[nodiscard]] int WideCharFlags() noexcept
    {
        return WC_ERR_INVALID_CHARS;
    }

    [[nodiscard]] int MultiByteFlags() noexcept
    {
        return MB_ERR_INVALID_CHARS;
    }

    [[nodiscard]] std::runtime_error MakeUtf8RuntimeError(std::string_view message)
    {
        return std::runtime_error(std::string(message));
    }
}

namespace pptxvp::helper
{
    std::wstring Utf8ToWide(std::string_view utf8)
    {
        if (utf8.empty())
        {
            return {};
        }

        int required_length = ::MultiByteToWideChar(
            CP_UTF8,
            MultiByteFlags(),
            utf8.data(),
            static_cast<int>(utf8.size()),
            nullptr,
            0);

        if (required_length <= 0)
        {
            throw MakeUtf8RuntimeError("UTF-8 文本转换为 UTF-16 失败。");
        }

        std::wstring wide(static_cast<std::size_t>(required_length), L'\0');

        int converted_length = ::MultiByteToWideChar(
            CP_UTF8,
            MultiByteFlags(),
            utf8.data(),
            static_cast<int>(utf8.size()),
            wide.data(),
            required_length);

        if (converted_length != required_length)
        {
            throw MakeUtf8RuntimeError("UTF-8 文本转换为 UTF-16 失败。");
        }

        return wide;
    }

    std::string WideToUtf8(std::wstring_view wide)
    {
        if (wide.empty())
        {
            return {};
        }

        int required_length = ::WideCharToMultiByte(
            CP_UTF8,
            WideCharFlags(),
            wide.data(),
            static_cast<int>(wide.size()),
            nullptr,
            0,
            nullptr,
            nullptr);

        if (required_length <= 0)
        {
            throw MakeUtf8RuntimeError("UTF-16 文本转换为 UTF-8 失败。");
        }

        std::string utf8(static_cast<std::size_t>(required_length), '\0');

        int converted_length = ::WideCharToMultiByte(
            CP_UTF8,
            WideCharFlags(),
            wide.data(),
            static_cast<int>(wide.size()),
            utf8.data(),
            required_length,
            nullptr,
            nullptr);

        if (converted_length != required_length)
        {
            throw MakeUtf8RuntimeError("UTF-16 文本转换为 UTF-8 失败。");
        }

        return utf8;
    }

    std::string StripUtf8Bom(std::string_view text)
    {
        constexpr std::string_view bom("\xEF\xBB\xBF", 3);

        if (text.starts_with(bom))
        {
            return std::string(text.substr(bom.size()));
        }

        return std::string(text);
    }

    std::wstring ToLowerAscii(std::wstring_view value)
    {
        std::wstring lowered;
        lowered.reserve(value.size());

        for (wchar_t character : value)
        {
            lowered.push_back(static_cast<wchar_t>(std::towlower(character)));
        }

        return lowered;
    }

    std::string ToLowerAscii(std::string_view value)
    {
        std::string lowered;
        lowered.reserve(value.size());

        for (unsigned char character : value)
        {
            lowered.push_back(static_cast<char>(std::tolower(character)));
        }

        return lowered;
    }
}
