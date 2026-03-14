export module PptxVideoProcessing.Helper.Console;

import std.compat;

export namespace pptxvp::helper
{
    void InitializeUtf16Console();
    void WriteLine(std::wstring_view message);
    void WriteErrorLine(std::wstring_view message);
    void WriteProgressLine(std::wstring_view message);
    void FinishProgressLine();
    void WaitForAnyKey(std::wstring_view prompt = L"按任意键退出...");
}
