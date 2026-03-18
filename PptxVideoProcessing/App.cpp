module;

module PptxVideoProcessing.App;

import std.compat;

import PptxVideoProcessing.Archive;
import PptxVideoProcessing.Config;
import PptxVideoProcessing.Helper.Console;
import PptxVideoProcessing.Helper.FileSystem;
import PptxVideoProcessing.Helper.Process;
import PptxVideoProcessing.Helper.Utf;
import PptxVideoProcessing.Media;
import PptxVideoProcessing.OfficeXml;
import PptxVideoProcessing.Ui;

namespace
{
    constexpr int SupportedFfmpegMajorVersion = 7;
    constexpr int SupportedFfmpegMaxMinorVersion = 1;
    constexpr int SupportedFfmpegLibavcodecMajor = 61;
    constexpr std::wstring_view SupportedFfmpegReleaseRange = L"FFmpeg 7.0.x / 7.1.x 稳定发行版";

    struct FfmpegBuildInfo
    {
        std::wstring version_line;
        std::optional<int> major_version;
        std::optional<int> minor_version;
        bool is_git_build{};
        std::optional<int> libavcodec_major;
    };

    [[nodiscard]] std::runtime_error MakeAppError(std::wstring_view message)
    {
        return std::runtime_error(pptxvp::helper::WideToUtf8(message));
    }

    [[nodiscard]] std::string_view TrimAsciiWhitespace(std::string_view text)
    {
        while (!text.empty() && std::isspace(static_cast<unsigned char>(text.front())))
        {
            text.remove_prefix(1);
        }

        while (!text.empty() && std::isspace(static_cast<unsigned char>(text.back())))
        {
            text.remove_suffix(1);
        }

        return text;
    }

    [[nodiscard]] std::optional<FfmpegBuildInfo> QueryFfmpegBuildInfo(const std::filesystem::path& ffmpeg_path)
    {
        const pptxvp::helper::ProcessResult result =
            pptxvp::helper::RunProcess(ffmpeg_path, {L"-version"}, ffmpeg_path.parent_path());

        if (result.exit_code != 0)
        {
            return std::nullopt;
        }

        FfmpegBuildInfo info;
        std::istringstream stream(result.output);
        std::string line;
        static const std::regex version_regex(R"(^ffmpeg version\s+(?:n)?(\d+)\.(\d+)(?:\.(\d+))?)", std::regex::icase);
        static const std::regex libavcodec_regex(R"(libavcodec\s+(\d+)\.\s*\d+\.\s*\d+)", std::regex::icase);
        std::smatch match;

        while (std::getline(stream, line))
        {
            const std::string trimmed_line(TrimAsciiWhitespace(line));

            if (trimmed_line.empty())
            {
                continue;
            }

            if (info.version_line.empty() &&
                trimmed_line.starts_with("ffmpeg version "))
            {
                info.version_line = pptxvp::helper::Utf8ToWide(trimmed_line);
                info.is_git_build = trimmed_line.find("git-") != std::string::npos;

                if (std::regex_search(trimmed_line, match, version_regex))
                {
                    info.major_version = std::stoi(match[1].str());
                    info.minor_version = std::stoi(match[2].str());
                }
            }

            if (!info.libavcodec_major.has_value() &&
                std::regex_search(trimmed_line, match, libavcodec_regex))
            {
                info.libavcodec_major = std::stoi(match[1].str());
            }
        }

        return info.version_line.empty() ? std::nullopt : std::optional<FfmpegBuildInfo>(std::move(info));
    }

    void ValidateFfmpegBuild(const std::filesystem::path& ffmpeg_path)
    {
        const std::optional<FfmpegBuildInfo> build_info = QueryFfmpegBuildInfo(ffmpeg_path);

        if (!build_info.has_value())
        {
            throw MakeAppError(
                L"无法识别 ffmpeg 版本信息。当前软件仅支持 " + std::wstring(SupportedFfmpegReleaseRange) +
                L"（libavcodec " + std::to_wstring(SupportedFfmpegLibavcodecMajor) + L".x）。");
        }

        if (build_info->is_git_build)
        {
            throw MakeAppError(
                L"检测到不受支持的 ffmpeg git/nightly build： " + build_info->version_line +
                L"。为避免 NVENC 驱动门槛被新 nightly 抬高，当前软件已锁定到 " +
                std::wstring(SupportedFfmpegReleaseRange) + L"（libavcodec " +
                std::to_wstring(SupportedFfmpegLibavcodecMajor) + L".x）。");
        }

        if (!build_info->major_version.has_value() || !build_info->minor_version.has_value())
        {
            throw MakeAppError(
                L"无法解析 ffmpeg 版本号： " + build_info->version_line + L"。当前软件仅支持 " +
                std::wstring(SupportedFfmpegReleaseRange) + L"（libavcodec " +
                std::to_wstring(SupportedFfmpegLibavcodecMajor) + L".x）。");
        }

        if (*build_info->major_version != SupportedFfmpegMajorVersion ||
            *build_info->minor_version > SupportedFfmpegMaxMinorVersion)
        {
            throw MakeAppError(
                L"当前 ffmpeg 版本不在支持范围内： " + build_info->version_line + L"。当前软件仅支持 " +
                std::wstring(SupportedFfmpegReleaseRange) + L"（libavcodec " +
                std::to_wstring(SupportedFfmpegLibavcodecMajor) + L".x）。");
        }

        if (build_info->libavcodec_major != SupportedFfmpegLibavcodecMajor)
        {
            const std::wstring detected_libavcodec =
                build_info->libavcodec_major.has_value()
                ? std::to_wstring(*build_info->libavcodec_major)
                : std::wstring(L"未知");
            throw MakeAppError(
                L"当前 ffmpeg 的 libavcodec 版本不在支持范围内： " + build_info->version_line +
                L"。检测到 libavcodec " + detected_libavcodec +
                L"，当前软件仅支持 libavcodec " +
                std::to_wstring(SupportedFfmpegLibavcodecMajor) + L".x。");
        }
    }

    [[nodiscard]] std::wstring StatusPrefix(pptxvp::MediaActionStatus status)
    {
        switch (status)
        {
        case pptxvp::MediaActionStatus::Processed:
            return L"[成功]";
        case pptxvp::MediaActionStatus::Failed:
            return L"[失败]";
        default:
            return L"[跳过]";
        }
    }

    void PrintSummary(
        const pptxvp::MediaProcessSummary& summary,
        const std::filesystem::path& input_path,
        const std::filesystem::path& output_path,
        bool copied_original)
    {
        pptxvp::helper::WriteLine(L"");
        pptxvp::helper::WriteLine(L"处理结果汇总");
        pptxvp::helper::WriteLine(L"成功处理： " + std::to_wstring(summary.processed_count));
        pptxvp::helper::WriteLine(L"跳过文件： " + std::to_wstring(summary.skipped_count));
        pptxvp::helper::WriteLine(L"处理失败： " + std::to_wstring(summary.failed_count));

        if (!summary.acceleration_backend.empty())
        {
            pptxvp::helper::WriteLine(L"加速后端： " + summary.acceleration_backend);
        }

        if (copied_original)
        {
            pptxvp::helper::WriteLine(L"由于没有媒体文件发生变更，本次直接复制了原始文件。");
        }

        for (const pptxvp::MediaActionResult& item : summary.items)
        {
            std::wstring line = StatusPrefix(item.status);
            line.append(L" ");
            line.append(item.source_path.filename().wstring());

            if (item.status == pptxvp::MediaActionStatus::Processed &&
                item.output_path.filename() != item.source_path.filename())
            {
                line.append(L" -> ");
                line.append(item.output_path.filename().wstring());
            }

            if (!item.message.empty())
            {
                line.append(L"：");
                line.append(item.message);
            }

            pptxvp::helper::WriteLine(line);
        }

        pptxvp::helper::WriteLine(L"源文件： " + input_path.wstring());

        if (!output_path.empty())
        {
            pptxvp::helper::WriteLine(L"输出文件： " + output_path.wstring());
        }
        else
        {
            pptxvp::helper::WriteLine(L"输出文件：本次未生成新的输出文件。");
        }
    }

    void EmitStage(
        const pptxvp::ProgressCallback& progress_callback,
        pptxvp::ProcessStage stage,
        std::wstring_view message,
        const std::filesystem::path& current_file = {},
        std::size_t current_index = 0,
        std::size_t total_count = 0,
        const std::optional<double>& file_percent = std::nullopt,
        std::wstring_view speed = {},
        std::wstring_view acceleration_backend = {})
    {
        if (!progress_callback)
        {
            if (!message.empty())
            {
                pptxvp::helper::WriteLine(message);
            }

            return;
        }

        progress_callback(pptxvp::ProgressEvent{
            .stage = stage,
            .current_file = current_file,
            .current_index = current_index,
            .total_count = total_count,
            .file_percent = file_percent,
            .message = std::wstring(message),
            .speed = std::wstring(speed),
            .acceleration_backend = std::wstring(acceleration_backend),
        });
    }

    [[nodiscard]] bool IsSupportedDirectMediaInputExtension(std::wstring_view extension)
    {
        static const std::array<std::wstring_view, 13> extensions = {
            L".mp4",
            L".m4v",
            L".mov",
            L".avi",
            L".wmv",
            L".mkv",
            L".webm",
            L".flv",
            L".mpeg",
            L".mpg",
            L".ts",
            L".mts",
            L".m2ts",
        };

        return std::ranges::find(extensions, extension) != extensions.end();
    }

    [[nodiscard]] std::filesystem::path BuildDirectMediaOutputPath(
        const std::filesystem::path& input_path,
        const std::filesystem::path& processed_path)
    {
        const std::filesystem::path desired = input_path.parent_path() /
            std::filesystem::path(input_path.stem().wstring() + L"_已处理" + processed_path.extension().wstring());
        return pptxvp::helper::MakeUniqueSiblingPath(desired);
    }

    [[nodiscard]] pptxvp::ProcessResult ProcessDirectMedia(
        const pptxvp::ProcessRequest& request,
        const pptxvp::AppConfig& config,
        const pptxvp::ProgressCallback& progress_callback)
    {
        pptxvp::helper::ScopedTempDirectory working_directory =
            pptxvp::helper::CreateUniqueTempDirectory(L"pptxvp-media");
        const std::filesystem::path media_directory = working_directory.path() / L"ppt" / L"media";
        std::filesystem::create_directories(media_directory);

        const std::filesystem::path staged_input = media_directory / request.input_path.filename();
        std::filesystem::copy_file(request.input_path, staged_input, std::filesystem::copy_options::overwrite_existing);

        EmitStage(progress_callback, pptxvp::ProcessStage::ProcessingMedia, L"正在处理视频...");
        pptxvp::MediaProcessSummary summary = pptxvp::ProcessMedia(
            working_directory.path(),
            request.ffmpeg_path,
            config,
            progress_callback
                ? pptxvp::MediaProgressCallback([&](const pptxvp::MediaProgressInfo& media_progress)
                    {
                        EmitStage(
                            progress_callback,
                            pptxvp::ProcessStage::ProcessingMedia,
                            media_progress.note.empty() ? L"正在处理视频..." : media_progress.note,
                            request.input_path,
                            media_progress.current_index,
                            media_progress.total_count,
                            media_progress.file_percent,
                            media_progress.speed,
                            media_progress.acceleration_backend);
                    })
                : pptxvp::MediaProgressCallback{});

        std::filesystem::path output_path;

        if (summary.AnyChanges())
        {
            const auto processed_item = std::ranges::find_if(
                summary.items,
                [](const pptxvp::MediaActionResult& item)
                {
                    return item.status == pptxvp::MediaActionStatus::Processed;
                });

            if (processed_item != summary.items.end())
            {
                output_path = BuildDirectMediaOutputPath(request.input_path, processed_item->output_path);
                std::filesystem::copy_file(processed_item->output_path, output_path);
                processed_item->output_path = output_path;
            }
        }

        const std::wstring acceleration_backend = summary.acceleration_backend;

        pptxvp::ProcessResult result{
            .input_path = request.input_path,
            .output_path = output_path,
            .summary = std::move(summary),
            .copied_original = false,
            .acceleration_backend = acceleration_backend,
        };

        if (!progress_callback)
        {
            PrintSummary(result.summary, result.input_path, result.output_path, result.copied_original);
        }

        EmitStage(
            progress_callback,
            pptxvp::ProcessStage::Completed,
            result.output_path.empty() ? L"未生成新的输出文件。" : L"处理完成。",
            result.output_path,
            0,
            0,
            std::nullopt,
            {},
            result.acceleration_backend);
        return result;
    }
}

namespace pptxvp
{
    ProcessResult ProcessPptx(const ProcessRequest& request, const ProgressCallback& progress_callback)
    {
        if (request.input_path.empty())
        {
            throw MakeAppError(L"未提供要处理的文件路径。");
        }

        if (!std::filesystem::exists(request.ffmpeg_path))
        {
            throw MakeAppError(
                L"未在程序同目录找到 ffmpeg.exe，期望路径为： " + request.ffmpeg_path.wstring());
        }

        ValidateFfmpegBuild(request.ffmpeg_path);

        EmitStage(progress_callback, ProcessStage::LoadingConfig, L"正在读取 config.json...");
        const AppConfig config = LoadConfig(request.config_path);

        EmitStage(progress_callback, ProcessStage::ValidatingInput, L"正在验证输入文件...");
        const std::wstring extension = helper::ToLowerAscii(request.input_path.extension().wstring());

        if (extension == L".ppt")
        {
            throw MakeAppError(L"当前版本仅支持 .pptx，不支持旧版二进制 .ppt 文件。");
        }

        const bool is_pptx_input = extension == L".pptx";
        const bool is_direct_media_input = IsSupportedDirectMediaInputExtension(extension);

        if (!is_pptx_input && !is_direct_media_input)
        {
            throw MakeAppError(L"所选文件不是受支持的 PPTX 或常见视频格式。");
        }

        if (!progress_callback)
        {
            helper::WriteLine(L"源文件： " + request.input_path.wstring());
        }

        if (!config.HasMediaChanges())
        {
            ProcessResult result{
                .input_path = request.input_path,
                .output_path = {},
                .summary = MediaProcessSummary{},
                .copied_original = false,
                .acceleration_backend = L"未启用",
            };

            if (!progress_callback)
            {
                PrintSummary(result.summary, result.input_path, result.output_path, result.copied_original);
            }

            EmitStage(progress_callback, ProcessStage::Completed, L"未生成新的输出文件。", {}, 0, 0, std::nullopt, {}, result.acceleration_backend);
            return result;
        }

        if (is_direct_media_input)
        {
            return ProcessDirectMedia(request, config, progress_callback);
        }

        helper::ScopedTempDirectory working_directory = helper::CreateUniqueTempDirectory(L"pptxvp");

        EmitStage(progress_callback, ProcessStage::ExtractingArchive, L"正在解包 PPTX...");
        ExtractArchive(request.input_path, working_directory.path());

        EmitStage(progress_callback, ProcessStage::ProcessingMedia, L"正在处理内嵌视频...");
        MediaProcessSummary summary = ProcessMedia(
            working_directory.path(),
            request.ffmpeg_path,
            config,
            progress_callback
                ? MediaProgressCallback([&](const MediaProgressInfo& media_progress)
                                        {
                                            EmitStage(
                                                progress_callback,
                                                ProcessStage::ProcessingMedia,
                                                media_progress.note.empty() ? L"正在处理内嵌视频..." : media_progress.note,
                                                media_progress.media_path,
                                                media_progress.current_index,
                                                media_progress.total_count,
                                                media_progress.file_percent,
                                                media_progress.speed,
                                                media_progress.acceleration_backend);
                                        })
                : MediaProgressCallback{});

        std::filesystem::path output_path;

        if (summary.AnyChanges())
        {
            output_path = helper::MakeUniqueProcessedOutputPath(request.input_path);

            if (!progress_callback)
            {
                helper::WriteLine(L"目标输出： " + output_path.wstring());
            }

            if (!summary.renames.empty())
            {
                EmitStage(progress_callback, ProcessStage::UpdatingReferences, L"正在更新 PPTX 内的媒体引用...");
                [[maybe_unused]] const std::size_t updated_reference_count =
                    UpdateOfficeMediaReferences(working_directory.path(), summary.renames);
            }

            EmitStage(progress_callback, ProcessStage::CreatingArchive, L"正在重新打包 PPTX...");
            CreateArchiveFromDirectory(working_directory.path(), output_path);
        }

        const std::wstring acceleration_backend = summary.acceleration_backend;

        ProcessResult result{
            .input_path = request.input_path,
            .output_path = output_path,
            .summary = std::move(summary),
            .copied_original = false,
            .acceleration_backend = acceleration_backend,
        };

        if (!progress_callback)
        {
            PrintSummary(result.summary, result.input_path, result.output_path, result.copied_original);
        }

        EmitStage(progress_callback, ProcessStage::Completed, result.output_path.empty() ? L"未生成新的输出文件。" : L"处理完成。", result.output_path, 0, 0, std::nullopt, {}, result.acceleration_backend);
        return result;
    }
    int Run()
    {
        const std::filesystem::path executable_directory = helper::GetExecutableDirectory();

        helper::WriteLine(L"请选择要处理的 PPTX 或视频文件...");
        const std::filesystem::path input_path = PickInputFile();

        if (input_path.empty())
        {
            helper::WriteLine(L"未选择文件，操作已取消。");
            return 0;
        }

        [[maybe_unused]] const ProcessResult result = ProcessPptx(ProcessRequest{
            .input_path = input_path,
            .ffmpeg_path = executable_directory / L"ffmpeg.exe",
            .config_path = executable_directory / L"config.json",
        });

        return 0;
    }
}







