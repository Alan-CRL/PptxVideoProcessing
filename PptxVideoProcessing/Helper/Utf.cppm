export module PptxVideoProcessing.Helper.Utf;

import std.compat;

export namespace pptxvp::helper
{
    [[nodiscard]] std::wstring Utf8ToWide(std::string_view utf8);
    [[nodiscard]] std::string WideToUtf8(std::wstring_view wide);
    [[nodiscard]] std::string StripUtf8Bom(std::string_view text);
    [[nodiscard]] std::wstring ToLowerAscii(std::wstring_view value);
    [[nodiscard]] std::string ToLowerAscii(std::string_view value);
}
