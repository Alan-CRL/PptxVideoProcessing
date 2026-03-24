/*
PptxVideoProcessing 使用说明

一、程序用途
1. 本程序用于处理 .pptx 演示文稿中的内嵌视频。
2. 本程序也支持直接处理常见视频文件。
2. 程序会解包 PPTX，查找 ppt/media 目录下的常见视频文件，并调用 ffmpeg.exe 进行转码或调整。
3. 处理完成后，程序会重新打包为新的 .pptx 文件，原文件不会被直接覆盖。

二、使用步骤
1. 将本程序生成的 exe 与 ffmpeg.exe 放在同一目录。
   - ffmpeg 需要 7.0 及以上版本；建议优先使用稳定发行版。
2. 可选：在 exe 同目录放置 config.json；如果不存在，程序首次启动时会自动创建一个默认包含 {"hardwareAcceleration":"auto"} 的配置。
3. 运行程序，按提示选择一个 .pptx 文件。
4. 程序会自动处理内嵌视频，并在原文件所在目录生成“原文件名_已处理.pptx”。
5. 如果输出文件已存在，程序会自动递增命名为“_已处理(2)”“_已处理(3)”等。

三、config.json 示例
{
  "encoder": "h264",
  "hardwareAcceleration": "auto",
  "presetLevel": "medium",
  "frameRate": 30,
  "resolution": "720p",
  "volumePercent": 100,
  "mute": false
}

四、config.json 配置项说明
1. encoder
   - 可选。
   - 用于指定视频编码器。
   - 支持别名：
	 h264  -> libx264
	 h265  -> libx265
	 av1   -> libsvtav1
	 mpeg4 -> mpeg4
   - 也支持直接填写 ffmpeg.exe 支持的编码器名称，例如：
	 libx264、libx265、libsvtav1、mpeg4、
	 h264_nvenc、hevc_nvenc、av1_nvenc、
	 h264_qsv、hevc_qsv、av1_qsv 等。
   - 如果该项缺省，而 frameRate 或 resolution 有值，程序会尽量探测原视频编码并保持原编码风格。

2. frameRate
   - 可选。
   - 必须为正整数。
   - 常见值：24、25、30、50、60。
   - 缺省表示不修改帧率。

3. hardwareAcceleration
   - 可选。
   - 用于优先选择 Windows 常见硬件编码器。
   - 当前支持：auto、none、nvidia、intel、amd、windows。
   - 缺省该项时，程序会按 auto 处理：自动识别并优先尝试当前机器真正可用的 NVIDIA / Intel / AMD / 系统原生编码；若都不可用，则回退为纯软件编码。
   - 若写 auto，程序会按当前机器和 ffmpeg 可用编码器自动选择 NVIDIA / Intel / AMD / 系统原生加速；若都不可用，则回退为纯软件编码。
   - 例如 encoder 为 h264 且 hardwareAcceleration 为 nvidia 时，会优先使用 h264_nvenc；若写 windows，则会优先尝试 h264_mf。
   - 如果你已经直接填写了 h264_nvenc、hevc_qsv 这类具体编码器名，则以 encoder 为准。

4. preset
   - 可选。
   - 用于手动填写具体编码预设值。
   - 例如软件编码常见值可用 veryfast、fast、medium；NVENC 常见值可用 p1 到 p7。
   - 若同时设置 preset 与 presetLevel，则以 preset 为准。

5. presetLevel
   - 可选。
   - 当前支持：low、medium、high。
   - 用于给普通用户快速选择常见预设档位。
   - 程序会按实际编码器自动映射，例如软件编码常用 fast / medium / slow，NVENC 常用 p3 / p4 / p5。

6. resolution
   - 可选。
   - 当前支持：360p、480p、720p、1080p、2160p。
   - 该项表示输出视频的高度上限，仅当原视频高度大于该值时才会缩小。
   - 缩放时会保持原视频宽高比。
   - 缺省表示不修改分辨率。

7. volumePercent
   - 可选。
   - 用于在原始音量基础上按 0% 到 300% 调整。
   - 100 表示保持原样，0 表示输出静音。

8. mute
   - 可选。
   - 布尔值，true 表示静音。
   - 若同时设置 mute=true 与 volumePercent，则以静音为准。

五、补充说明
1. 仅支持 .pptx，不支持旧版二进制 .ppt。
2. config.json 支持 UTF-8 和 UTF-8 BOM。
3. 某个配置项缺省时，仅表示不修改该项，不会报错。
4. 程序会扫描 ppt/media 中的常见视频文件，例如 mp4、m4v、mov、avi、wmv、mkv、webm。
5. 若原视频容器与目标编码不兼容，程序会转成 MP4，并自动更新 PPTX 内的媒体引用。
6. 程序默认保留首个视频流和可选音频流，音频通常按复制方式保留，不主动改变演示稿中的媒体位置与时长设置。
7. 若检测到 ffmpeg git/nightly build，程序会给出警告，但不会禁止继续处理；建议优先使用稳定发行版，以减少 NVENC / AMF / QSV 驱动门槛波动带来的风险。
*/
#include <nlohmann/json.hpp>

import std.compat;

import PptxVideoProcessing.App;
import PptxVideoProcessing.Helper.Console;
import PptxVideoProcessing.Helper.FileSystem;
import PptxVideoProcessing.Helper.Utf;

namespace
{
	using json = nlohmann::json;

	struct WorkerOptions
	{
		std::optional<std::filesystem::path> input_path;
		bool json_progress{};
		bool no_pause{};
		bool show_help{};
	};

	[[nodiscard]] std::runtime_error MakeArgumentError(std::wstring_view message)
	{
		return std::runtime_error(pptxvp::helper::WideToUtf8(message));
	}

	[[nodiscard]] std::string PathToUtf8(const std::filesystem::path& path)
	{
		return pptxvp::helper::WideToUtf8(path.wstring());
	}

	[[nodiscard]] std::string StageToCode(pptxvp::ProcessStage stage)
	{
		switch (stage)
		{
		case pptxvp::ProcessStage::LoadingConfig:
			return "loading_config";
		case pptxvp::ProcessStage::ValidatingInput:
			return "validating_input";
		case pptxvp::ProcessStage::CopyingOriginal:
			return "copying_original";
		case pptxvp::ProcessStage::ExtractingArchive:
			return "extracting_archive";
		case pptxvp::ProcessStage::ProcessingMedia:
			return "processing_media";
		case pptxvp::ProcessStage::UpdatingReferences:
			return "updating_references";
		case pptxvp::ProcessStage::CreatingArchive:
			return "creating_archive";
		case pptxvp::ProcessStage::Completed:
			return "completed";
		}

		return "unknown";
	}

	void WriteJsonLine(const json& value)
	{
		std::cout << value.dump() << '\n';
		std::cout.flush();
	}

	void PrintBanner()
	{
		pptxvp::helper::WriteLine(L"版权所有 (c) 2026 AlanCRL(陈润林) 工作室");
		pptxvp::helper::WriteLine(L"本项目基于 GNU 通用公共许可证第 3 版获得许可");
		pptxvp::helper::WriteLine(L"------------------------------------------------");
		pptxvp::helper::WriteLine(L"");
	}

	void PrintUsage()
	{
		pptxvp::helper::WriteLine(L"用法：");
		pptxvp::helper::WriteLine(L"  PptxVideoProcessing.Worker.exe");
		pptxvp::helper::WriteLine(L"  PptxVideoProcessing.Worker.exe --input <文件路径> [--no-pause]");
		pptxvp::helper::WriteLine(L"  PptxVideoProcessing.Worker.exe --input <文件路径> --json-progress --no-pause");
		pptxvp::helper::WriteLine(L"");
		pptxvp::helper::WriteLine(L"参数说明：");
		pptxvp::helper::WriteLine(L"  --input <路径>        处理指定的 PPTX 或视频文件");
		pptxvp::helper::WriteLine(L"  --json-progress       以 JSON Lines 输出进度事件");
		pptxvp::helper::WriteLine(L"  --no-pause            结束时不等待按键");
		pptxvp::helper::WriteLine(L"  --help                显示帮助");
	}

	[[nodiscard]] WorkerOptions ParseOptions(int argc, wchar_t* argv[])
	{
		WorkerOptions options;

		for (int index = 1; index < argc; ++index)
		{
			const std::wstring_view argument = argv[index];

			if (argument == L"--input")
			{
				if (index + 1 >= argc)
				{
					throw MakeArgumentError(L"--input 需要提供一个文件路径。");
				}

				options.input_path = std::filesystem::path(argv[++index]);
				continue;
			}

			if (argument == L"--json-progress")
			{
				options.json_progress = true;
				options.no_pause = true;
				continue;
			}

			if (argument == L"--no-pause")
			{
				options.no_pause = true;
				continue;
			}

			if (argument == L"--help" || argument == L"-h" || argument == L"/?")
			{
				options.show_help = true;
				continue;
			}

			throw MakeArgumentError(L"不支持的参数： " + std::wstring(argument));
		}

		if (options.json_progress && !options.input_path.has_value())
		{
			throw MakeArgumentError(L"--json-progress 模式必须同时提供 --input。");
		}

		return options;
	}

	json SerializeProgressEvent(const pptxvp::ProgressEvent& event)
	{
		json payload = {
			{"type", "progress"},
			{"stage", StageToCode(event.stage)},
			{"message", pptxvp::helper::WideToUtf8(event.message)},
			{"currentIndex", event.current_index},
			{"totalCount", event.total_count},
			{"speed", pptxvp::helper::WideToUtf8(event.speed)},
		};

		if (!event.acceleration_backend.empty())
		{
			payload["accelerationBackend"] = pptxvp::helper::WideToUtf8(event.acceleration_backend);
		}

		if (!event.current_file.empty())
		{
			payload["currentFile"] = PathToUtf8(event.current_file);
		}

		if (event.file_percent.has_value())
		{
			payload["filePercent"] = *event.file_percent;
		}

		return payload;
	}

	json SerializeCompletedEvent(const pptxvp::ProcessResult& result)
	{
		json payload = {
			{"type", "completed"},
			{"inputPath", PathToUtf8(result.input_path)},
			{"outputPath", PathToUtf8(result.output_path)},
			{"copiedOriginal", result.copied_original},
			{"processedCount", result.summary.processed_count},
			{"skippedCount", result.summary.skipped_count},
			{"failedCount", result.summary.failed_count},
			{"alreadySatisfiedCount", result.summary.already_satisfied_count},
			{"noVideoCount", result.summary.no_video_count},
		};

		if (!result.acceleration_backend.empty())
		{
			payload["accelerationBackend"] = pptxvp::helper::WideToUtf8(result.acceleration_backend);
		}

		return payload;
	}

	void WriteJsonError(std::string_view message)
	{
		WriteJsonLine(json{
			{"type", "error"},
			{"message", std::string(message)},
			});
	}
}

int wmain(int argc, wchar_t* argv[])
{
	WorkerOptions options;

	try
	{
		options = ParseOptions(argc, argv);
	}
	catch (const std::exception& exception)
	{
		pptxvp::helper::InitializeUtf16Console();
		pptxvp::helper::WriteErrorLine(L"错误： " + pptxvp::helper::Utf8ToWide(exception.what()));
		return 1;
	}

	const bool interactive_console = !options.json_progress;

	if (interactive_console)
	{
		pptxvp::helper::InitializeUtf16Console();
		PrintBanner();
	}

	if (options.show_help)
	{
		PrintUsage();
		return 0;
	}

	int exit_code = 0;

	try
	{
		if (!options.input_path.has_value())
		{
			exit_code = pptxvp::Run();
		}
		else
		{
			const std::filesystem::path executable_directory = pptxvp::helper::GetExecutableDirectory();
			const pptxvp::ProcessRequest request{
				.input_path = *options.input_path,
				.ffmpeg_path = executable_directory / L"ffmpeg.exe",
				.config_path = executable_directory / L"config.json",
			};

			const pptxvp::ProcessResult result = pptxvp::ProcessPptx(
				request,
				options.json_progress
				? pptxvp::ProgressCallback([](const pptxvp::ProgressEvent& event)
					{
						WriteJsonLine(SerializeProgressEvent(event));
					})
				: pptxvp::ProgressCallback{});

			if (options.json_progress)
			{
				WriteJsonLine(SerializeCompletedEvent(result));
			}
		}
	}
	catch (const std::exception& exception)
	{
		if (options.json_progress)
		{
			WriteJsonError(exception.what());
		}
		else
		{
			pptxvp::helper::WriteErrorLine(L"错误： " + pptxvp::helper::Utf8ToWide(exception.what()));
		}

		exit_code = 1;
	}
	catch (...)
	{
		if (options.json_progress)
		{
			WriteJsonError("发生了未预期的异常。");
		}
		else
		{
			pptxvp::helper::WriteErrorLine(L"错误：发生了未预期的异常。");
		}

		exit_code = 1;
	}

	if (interactive_console && !options.no_pause)
	{
		pptxvp::helper::WriteLine(L"");
		pptxvp::helper::WaitForAnyKey();
	}

	return exit_code;
}

