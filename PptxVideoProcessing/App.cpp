module;

module PptxVideoProcessing.App;

import std.compat;

import PptxVideoProcessing.Archive;
import PptxVideoProcessing.Config;
import PptxVideoProcessing.Helper.Console;
import PptxVideoProcessing.Helper.FileSystem;
import PptxVideoProcessing.Helper.Utf;
import PptxVideoProcessing.Media;
import PptxVideoProcessing.OfficeXml;
import PptxVideoProcessing.Ui;

namespace
{
    [[nodiscard]] std::runtime_error MakeAppError(std::wstring_view message)
    {
        return std::runtime_error(pptxvp::helper::WideToUtf8(message));
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
            pptxvp::helper::WriteLine(L"由于没有媒体文件发生变更，本次直接复制了原始 PPTX。");
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
            pptxvp::helper::WriteLine(L"输出文件：本次未生成新的 PPTX。");
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
}

namespace pptxvp
{
    ProcessResult ProcessPptx(const ProcessRequest& request, const ProgressCallback& progress_callback)
    {
        if (request.input_path.empty())
        {
            throw MakeAppError(L"未提供要处理的 PPTX 文件路径。");
        }

        if (!std::filesystem::exists(request.ffmpeg_path))
        {
            throw MakeAppError(
                L"未在程序同目录找到 ffmpeg.exe，期望路径为： " + request.ffmpeg_path.wstring());
        }

        EmitStage(progress_callback, ProcessStage::LoadingConfig, L"正在读取 config.json...");
        const AppConfig config = LoadConfig(request.config_path);

        EmitStage(progress_callback, ProcessStage::ValidatingInput, L"正在验证输入文件...");
        const std::wstring extension = helper::ToLowerAscii(request.input_path.extension().wstring());

        if (extension == L".ppt")
        {
            throw MakeAppError(L"当前版本仅支持 .pptx，不支持旧版二进制 .ppt 文件。");
        }

        if (extension != L".pptx")
        {
            throw MakeAppError(L"所选文件不是有效的 .pptx 演示文稿包。");
        }

        if (!progress_callback)
        {
            helper::WriteLine(L"源文件： " + request.input_path.wstring());
        }

        if (!config.HasVideoChanges())
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

        helper::WriteLine(L"请选择要处理的 PPTX 文件...");
        const std::filesystem::path input_path = PickPowerPointFile();

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







