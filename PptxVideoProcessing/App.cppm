export module PptxVideoProcessing.App;

import std.compat;
import PptxVideoProcessing.Media;

export namespace pptxvp
{
    enum class ProcessStage
    {
        LoadingConfig,
        ValidatingInput,
        CopyingOriginal,
        ExtractingArchive,
        ProcessingMedia,
        UpdatingReferences,
        CreatingArchive,
        Completed,
    };

    struct ProcessRequest
    {
        std::filesystem::path input_path;
        std::filesystem::path ffmpeg_path;
        std::filesystem::path config_path;
    };

    struct ProgressEvent
    {
        ProcessStage stage{ProcessStage::LoadingConfig};
        std::filesystem::path current_file;
        std::size_t current_index{};
        std::size_t total_count{};
        std::optional<double> file_percent;
        std::wstring message;
        std::wstring speed;
        std::wstring acceleration_backend;
    };

    using ProgressCallback = std::function<void(const ProgressEvent&)>;

    struct ProcessResult
    {
        std::filesystem::path input_path;
        std::filesystem::path output_path;
        MediaProcessSummary summary;
        bool copied_original{};
        std::wstring acceleration_backend;
    };

    [[nodiscard]] ProcessResult ProcessPptx(
        const ProcessRequest& request,
        const ProgressCallback& progress_callback = {});

    [[nodiscard]] int Run();
}
