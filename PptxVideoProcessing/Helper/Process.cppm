export module PptxVideoProcessing.Helper.Process;

import std.compat;

export namespace pptxvp::helper
{
    struct ProcessResult
    {
        int exit_code{};
        std::string output;
    };

    using OutputChunkCallback = std::function<void(std::string_view)>;

    [[nodiscard]] ProcessResult RunProcess(
        const std::filesystem::path& executable_path,
        const std::vector<std::wstring>& arguments,
        const std::filesystem::path& working_directory = {});

    [[nodiscard]] ProcessResult RunProcessStreaming(
        const std::filesystem::path& executable_path,
        const std::vector<std::wstring>& arguments,
        const OutputChunkCallback& output_callback,
        const std::filesystem::path& working_directory = {});
}
