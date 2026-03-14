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
        pptxvp::helper::WriteLine(L"输出文件： " + output_path.wstring());
    }
}

namespace pptxvp
{
    int Run()
    {
        const std::filesystem::path executable_directory = helper::GetExecutableDirectory();
        const std::filesystem::path ffmpeg_path = executable_directory / L"ffmpeg.exe";
        const std::filesystem::path config_path = executable_directory / L"config.json";

        if (!std::filesystem::exists(ffmpeg_path))
        {
            throw MakeAppError(
                L"未在程序同目录找到 ffmpeg.exe，期望路径为： " + ffmpeg_path.wstring());
        }

        helper::WriteLine(L"正在读取 config.json...");
        const AppConfig config = LoadConfig(config_path);

        helper::WriteLine(L"请选择要处理的 PPTX 文件...");
        const std::filesystem::path input_path = PickPowerPointFile();

        if (input_path.empty())
        {
            helper::WriteLine(L"未选择文件，操作已取消。");
            return 0;
        }

        const std::wstring extension = helper::ToLowerAscii(input_path.extension().wstring());

        if (extension == L".ppt")
        {
            throw MakeAppError(L"当前版本仅支持 .pptx，不支持旧版二进制 .ppt 文件。");
        }

        if (extension != L".pptx")
        {
            throw MakeAppError(L"所选文件不是有效的 .pptx 演示文稿包。");
        }

        const std::filesystem::path output_path = helper::MakeUniqueProcessedOutputPath(input_path);
        helper::WriteLine(L"源文件： " + input_path.wstring());
        helper::WriteLine(L"目标输出： " + output_path.wstring());

        if (!config.HasVideoChanges())
        {
            std::filesystem::copy_file(input_path, output_path);
            PrintSummary(MediaProcessSummary{}, input_path, output_path, true);
            return 0;
        }

        helper::ScopedTempDirectory working_directory = helper::CreateUniqueTempDirectory(L"pptxvp");

        helper::WriteLine(L"正在解包 PPTX...");
        ExtractArchive(input_path, working_directory.path());

        helper::WriteLine(L"正在处理内嵌视频...");
        MediaProcessSummary summary = ProcessMedia(working_directory.path(), ffmpeg_path, config);

        bool copied_original = false;

        if (!summary.AnyChanges())
        {
            std::filesystem::copy_file(input_path, output_path);
            copied_original = true;
        }
        else
        {
            if (!summary.renames.empty())
            {
                helper::WriteLine(L"正在更新 PPTX 内的媒体引用...");
                [[maybe_unused]] const std::size_t updated_reference_count =
                    UpdateOfficeMediaReferences(working_directory.path(), summary.renames);
            }

            helper::WriteLine(L"正在重新打包 PPTX...");
            CreateArchiveFromDirectory(working_directory.path(), output_path);
        }

        PrintSummary(summary, input_path, output_path, copied_original);
        return 0;
    }
}
