export module PptxVideoProcessing.Media;

import std.compat;

import PptxVideoProcessing.Config;

export namespace pptxvp
{
    enum class MediaActionStatus
    {
        Processed,
        Skipped,
        Failed,
    };

    struct MediaRename
    {
        std::filesystem::path original_relative_path;
        std::filesystem::path updated_relative_path;
    };

    struct MediaActionResult
    {
        std::filesystem::path source_path;
        std::filesystem::path output_path;
        MediaActionStatus status{MediaActionStatus::Skipped};
        std::wstring message;
    };

    struct MediaProcessSummary
    {
        std::vector<MediaActionResult> items;
        std::vector<MediaRename> renames;
        std::size_t processed_count{};
        std::size_t skipped_count{};
        std::size_t failed_count{};
        std::size_t already_satisfied_count{};
        std::size_t no_video_count{};
        std::wstring acceleration_backend;

        [[nodiscard]] bool AnyChanges() const noexcept
        {
            return processed_count != 0U;
        }
    };

    struct MediaProgressInfo
    {
        std::filesystem::path media_path;
        std::size_t current_index{};
        std::size_t total_count{};
        std::optional<double> current_seconds;
        std::optional<double> total_seconds;
        std::optional<double> file_percent;
        std::wstring note;
        std::wstring speed;
        std::wstring acceleration_backend;
    };

    using MediaProgressCallback = std::function<void(const MediaProgressInfo&)>;

    [[nodiscard]] MediaProcessSummary ProcessMedia(
        const std::filesystem::path& extracted_root,
        const std::filesystem::path& ffmpeg_path,
        const AppConfig& config,
        const MediaProgressCallback& progress_callback = {});
}
