export module PptxVideoProcessing.Archive;

import std.compat;

export namespace pptxvp
{
    void ExtractArchive(const std::filesystem::path& archive_path, const std::filesystem::path& destination_directory);
    void CreateArchiveFromDirectory(
        const std::filesystem::path& source_directory,
        const std::filesystem::path& archive_path);
}
