export module PptxVideoProcessing.Helper.FileSystem;

import std.compat;

export namespace pptxvp::helper
{
    class ScopedTempDirectory
    {
    public:
        ScopedTempDirectory() = default;
        explicit ScopedTempDirectory(std::filesystem::path path);
        ScopedTempDirectory(ScopedTempDirectory&& other) noexcept;
        ScopedTempDirectory& operator=(ScopedTempDirectory&& other) noexcept;
        ScopedTempDirectory(const ScopedTempDirectory&) = delete;
        ScopedTempDirectory& operator=(const ScopedTempDirectory&) = delete;
        ~ScopedTempDirectory();

        [[nodiscard]] const std::filesystem::path& path() const noexcept;
        [[nodiscard]] bool empty() const noexcept;
        void reset();

    private:
        std::filesystem::path path_{};
    };

    [[nodiscard]] std::filesystem::path GetExecutablePath();
    [[nodiscard]] std::filesystem::path GetExecutableDirectory();
    [[nodiscard]] ScopedTempDirectory CreateUniqueTempDirectory(std::wstring_view prefix);
    [[nodiscard]] std::string ReadTextFileUtf8(const std::filesystem::path& path);
    void WriteTextFileUtf8(const std::filesystem::path& path, std::string_view text);
    [[nodiscard]] std::filesystem::path MakeUniqueProcessedOutputPath(const std::filesystem::path& input_path);
    [[nodiscard]] std::filesystem::path MakeUniqueSiblingPath(const std::filesystem::path& desired_path);
}
