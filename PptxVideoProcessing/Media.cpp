module;

module PptxVideoProcessing.Media;

import std.compat;

import PptxVideoProcessing.Config;
import PptxVideoProcessing.Helper.Console;
import PptxVideoProcessing.Helper.FileSystem;
import PptxVideoProcessing.Helper.Process;
import PptxVideoProcessing.Helper.Utf;

namespace
{
    enum class CodecFamily
    {
        Unknown,
        H264,
        H265,
        Av1,
        Mpeg4,
    };

    struct ProbeInfo
    {
        bool has_video{};
        std::wstring codec_name;
        std::optional<double> duration_seconds;
    };

    struct TranscodeAttempt
    {
        bool succeeded{};
        std::filesystem::path final_path;
        std::wstring detail;
    };

    struct LiveProgressState
    {
        std::optional<double> current_seconds;
        std::wstring speed;
        bool finished{};
    };

    using LiveProgressCallback = std::function<void(const LiveProgressState&)>;

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

    [[nodiscard]] std::wstring AsciiToWide(std::string_view text)
    {
        return std::wstring(text.begin(), text.end());
    }

    [[nodiscard]] std::optional<double> ParseTimestampSeconds(std::string_view text)
    {
        text = TrimAsciiWhitespace(text);

        const std::size_t first_colon = text.find(':');
        const std::size_t second_colon = text.find(':', first_colon == std::string_view::npos ? 0 : first_colon + 1);

        if (first_colon == std::string_view::npos || second_colon == std::string_view::npos)
        {
            return std::nullopt;
        }

        try
        {
            const int hours = std::stoi(std::string(text.substr(0, first_colon)));
            const int minutes = std::stoi(std::string(text.substr(first_colon + 1, second_colon - first_colon - 1)));
            const double seconds = std::stod(std::string(text.substr(second_colon + 1)));
            return static_cast<double>(hours * 3600 + minutes * 60) + seconds;
        }
        catch (...)
        {
            return std::nullopt;
        }
    }

    [[nodiscard]] std::wstring FormatHmsFromSeconds(double seconds_value)
    {
        const auto total_seconds = static_cast<long long>(std::max(0.0, seconds_value));
        const long long hours = total_seconds / 3600;
        const long long minutes = (total_seconds % 3600) / 60;
        const long long seconds = total_seconds % 60;

        std::wostringstream stream;
        stream << std::setfill(L'0') << std::setw(2) << hours << L':' << std::setw(2) << minutes << L':' << std::setw(2)
               << seconds;
        return stream.str();
    }

    [[nodiscard]] std::wstring FormatElapsed(std::chrono::steady_clock::duration duration)
    {
        const auto seconds = std::chrono::duration_cast<std::chrono::seconds>(duration).count();
        return FormatHmsFromSeconds(static_cast<double>(seconds));
    }

    [[nodiscard]] std::wstring FormatPercent(double value)
    {
        std::wostringstream stream;
        stream << std::fixed << std::setprecision(1) << value;
        return stream.str();
    }

    [[nodiscard]] std::wstring BuildProgressLine(
        std::size_t current_index,
        std::size_t total_count,
        std::wstring_view file_name,
        std::wstring_view note,
        std::optional<double> current_seconds,
        std::optional<double> total_seconds,
        std::wstring_view speed,
        std::chrono::steady_clock::duration elapsed)
    {
        double file_ratio = 0.0;

        if (current_seconds.has_value() && total_seconds.has_value() && *total_seconds > 0.0)
        {
            file_ratio = std::clamp(*current_seconds / *total_seconds, 0.0, 1.0);
        }

        std::wstring line = L"总进度: ";
        line.append(std::to_wstring(current_index));
        line.append(L"/");
        line.append(std::to_wstring(total_count));
        line.append(L" | ");
        line.append(file_name);

        if (current_seconds.has_value())
        {
            line.append(L" | ");

            if (total_seconds.has_value() && *total_seconds > 0.0)
            {
                line.append(FormatPercent(file_ratio * 100.0));
                line.append(L"% (");
                line.append(FormatHmsFromSeconds(*current_seconds));
                line.append(L" / ");
                line.append(FormatHmsFromSeconds(*total_seconds));
                line.append(L")");
            }
            else
            {
                line.append(FormatHmsFromSeconds(*current_seconds));
            }
        }
        else if (!note.empty())
        {
            line.append(L" | ");
            line.append(note);
        }

        line.append(L" | 耗时 ");
        line.append(FormatElapsed(elapsed));

        if (!speed.empty())
        {
            line.append(L" | ");
            line.append(speed);
        }

        return line;
    }

    [[nodiscard]] bool IsFfmpegProgressLine(std::string_view line)
    {
        line = TrimAsciiWhitespace(line);

        return line.starts_with("frame=") || line.starts_with("fps=") || line.starts_with("stream_") ||
               line.starts_with("bitrate=") || line.starts_with("total_size=") || line.starts_with("out_time=") ||
               line.starts_with("out_time_ms=") || line.starts_with("out_time_us=") ||
               line.starts_with("dup_frames=") || line.starts_with("drop_frames=") || line.starts_with("speed=") ||
               line.starts_with("progress=");
    }

    class FfmpegProgressParser
    {
    public:
        void Consume(std::string_view chunk, const LiveProgressCallback& callback)
        {
            buffer_.append(chunk);

            while (true)
            {
                const std::size_t line_end = buffer_.find('\n');

                if (line_end == std::string::npos)
                {
                    break;
                }

                std::string line = buffer_.substr(0, line_end);
                buffer_.erase(0, line_end + 1);

                if (!line.empty() && line.back() == '\r')
                {
                    line.pop_back();
                }

                ProcessLine(line, callback);
            }
        }

    private:
        void ProcessLine(std::string_view line, const LiveProgressCallback& callback)
        {
            line = TrimAsciiWhitespace(line);

            if (line.empty())
            {
                return;
            }

            if (line.starts_with("out_time="))
            {
                state_.current_seconds = ParseTimestampSeconds(line.substr(std::string_view("out_time=").size()));
                return;
            }

            if (line.starts_with("speed="))
            {
                state_.speed = AsciiToWide(line.substr(std::string_view("speed=").size()));
                return;
            }

            if (line.starts_with("progress="))
            {
                state_.finished = line.substr(std::string_view("progress=").size()) == "end";

                if (callback)
                {
                    callback(state_);
                }
            }
        }

        std::string buffer_;
        LiveProgressState state_{};
    };

    [[nodiscard]] std::wstring LowerExtension(const std::filesystem::path& path)
    {
        return pptxvp::helper::ToLowerAscii(path.extension().wstring());
    }

    [[nodiscard]] bool IsCandidateMediaExtension(const std::filesystem::path& path)
    {
        static const std::array<std::wstring_view, 7> extensions = {
            L".mp4",
            L".m4v",
            L".mov",
            L".avi",
            L".wmv",
            L".mkv",
            L".webm",
        };

        const std::wstring extension = LowerExtension(path);
        return std::ranges::find(extensions, extension) != extensions.end();
    }

    [[nodiscard]] CodecFamily CodecFamilyFromDecoder(std::wstring_view codec)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(codec);

        if (lowered == L"h264")
        {
            return CodecFamily::H264;
        }

        if (lowered == L"hevc" || lowered == L"h265")
        {
            return CodecFamily::H265;
        }

        if (lowered == L"av1")
        {
            return CodecFamily::Av1;
        }

        if (lowered == L"mpeg4")
        {
            return CodecFamily::Mpeg4;
        }

        return CodecFamily::Unknown;
    }

    [[nodiscard]] CodecFamily CodecFamilyFromEncoder(std::wstring_view encoder)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(encoder);

        if (lowered.contains(L"x264") || lowered.starts_with(L"h264"))
        {
            return CodecFamily::H264;
        }

        if (lowered.contains(L"x265") || lowered.starts_with(L"hevc") || lowered == L"libx265")
        {
            return CodecFamily::H265;
        }

        if (lowered == L"mpeg4")
        {
            return CodecFamily::Mpeg4;
        }

        if (lowered.contains(L"av1") || lowered == L"librav1e" || lowered == L"libsvtav1")
        {
            return CodecFamily::Av1;
        }

        return CodecFamily::Unknown;
    }

    [[nodiscard]] std::optional<std::wstring> EncoderForCodecFamily(CodecFamily family)
    {
        switch (family)
        {
        case CodecFamily::H264:
            return L"libx264";
        case CodecFamily::H265:
            return L"libx265";
        case CodecFamily::Av1:
            return L"libsvtav1";
        case CodecFamily::Mpeg4:
            return L"mpeg4";
        default:
            return std::nullopt;
        }
    }

    [[nodiscard]] bool ContainerSupportsCodec(const std::filesystem::path& path, CodecFamily family)
    {
        const std::wstring extension = LowerExtension(path);

        if (extension == L".mkv")
        {
            return true;
        }

        if (extension == L".mp4" || extension == L".m4v")
        {
            return family == CodecFamily::H264 || family == CodecFamily::H265 || family == CodecFamily::Av1 ||
                   family == CodecFamily::Mpeg4;
        }

        if (extension == L".mov")
        {
            return family == CodecFamily::H264 || family == CodecFamily::H265 || family == CodecFamily::Mpeg4;
        }

        if (extension == L".webm")
        {
            return family == CodecFamily::Av1;
        }

        if (extension == L".avi")
        {
            return family == CodecFamily::Mpeg4;
        }

        return false;
    }

    [[nodiscard]] std::vector<std::filesystem::path> FindCandidateMediaFiles(const std::filesystem::path& media_directory)
    {
        std::vector<std::filesystem::path> files;

        if (!std::filesystem::exists(media_directory))
        {
            return files;
        }

        for (const auto& entry : std::filesystem::recursive_directory_iterator(media_directory))
        {
            if (entry.is_regular_file() && IsCandidateMediaExtension(entry.path()))
            {
                files.push_back(entry.path());
            }
        }

        std::ranges::sort(files);
        return files;
    }

    [[nodiscard]] ProbeInfo ProbeMedia(
        const std::filesystem::path& ffmpeg_path,
        const std::filesystem::path& media_path)
    {
        const pptxvp::helper::ProcessResult result =
            pptxvp::helper::RunProcess(ffmpeg_path, {L"-hide_banner", L"-i", media_path.wstring()});

        static const std::regex video_regex(R"(Video:\s*([A-Za-z0-9_]+))");
        static const std::regex duration_regex(R"(Duration:\s*([0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?))");
        std::smatch match;
        std::optional<double> duration_seconds;

        if (std::regex_search(result.output, match, duration_regex))
        {
            duration_seconds = ParseTimestampSeconds(match[1].str());
        }

        if (std::regex_search(result.output, match, video_regex))
        {
            return ProbeInfo{
                .has_video = true,
                .codec_name = pptxvp::helper::Utf8ToWide(match[1].str()),
                .duration_seconds = duration_seconds,
            };
        }

        return ProbeInfo{};
    }

    [[nodiscard]] std::optional<std::wstring> ResolveEffectiveEncoder(const pptxvp::AppConfig& config, const ProbeInfo& probe)
    {
        if (config.encoder.has_value())
        {
            return config.encoder;
        }

        if (!config.frame_rate.has_value() && !config.resolution_height.has_value())
        {
            return std::nullopt;
        }

        return EncoderForCodecFamily(CodecFamilyFromDecoder(probe.codec_name));
    }

    [[nodiscard]] std::wstring BuildFilterChain(const pptxvp::AppConfig& config)
    {
        std::vector<std::wstring> filters;

        if (config.frame_rate.has_value())
        {
            filters.push_back(L"fps=" + std::to_wstring(*config.frame_rate));
        }

        if (config.resolution_height.has_value())
        {
            filters.push_back(L"scale=-2:" + std::to_wstring(*config.resolution_height));
        }

        std::wstring chain;

        for (std::size_t index = 0; index < filters.size(); ++index)
        {
            if (index != 0)
            {
                chain.append(L",");
            }

            chain.append(filters[index]);
        }

        return chain;
    }

    [[nodiscard]] std::filesystem::path MakeConvertedTargetPath(const std::filesystem::path& media_path)
    {
        const std::filesystem::path desired =
            media_path.parent_path() / std::filesystem::path(media_path.stem().wstring() + L".mp4");
        return pptxvp::helper::MakeUniqueSiblingPath(desired);
    }

    [[nodiscard]] std::filesystem::path MakeTemporaryOutputPath(const std::filesystem::path& final_path)
    {
        const std::filesystem::path desired = final_path.parent_path() /
                                              std::filesystem::path(
                                                  final_path.stem().wstring() + L".pptxvp.tmp" + final_path.extension().wstring());
        return pptxvp::helper::MakeUniqueSiblingPath(desired);
    }

    [[nodiscard]] std::wstring ShortFailureDetail(std::string_view ffmpeg_output)
    {
        std::istringstream stream{std::string(ffmpeg_output)};
        std::string line;
        std::string last_non_empty_line;

        while (std::getline(stream, line))
        {
            const std::string_view trimmed_line = TrimAsciiWhitespace(line);

            if (!trimmed_line.empty() && !IsFfmpegProgressLine(trimmed_line))
            {
                last_non_empty_line = std::string(trimmed_line);
            }
        }

        if (last_non_empty_line.empty())
        {
            return L"ffmpeg 执行失败。";
        }

        return L"ffmpeg 执行失败： " + pptxvp::helper::Utf8ToWide(last_non_empty_line);
    }

    [[nodiscard]] std::vector<std::wstring> BuildTranscodeArguments(
        const std::filesystem::path& input_path,
        const std::filesystem::path& output_path,
        std::wstring_view encoder,
        const pptxvp::AppConfig& config)
    {
        std::vector<std::wstring> arguments = {
            L"-y",
            L"-hide_banner",
            L"-nostats",
            L"-progress",
            L"pipe:1",
            L"-i",
            input_path.wstring(),
            L"-map",
            L"0:v:0",
            L"-map",
            L"0:a:0?",
            L"-c:v",
            std::wstring(encoder),
        };

        const std::wstring filter_chain = BuildFilterChain(config);

        if (!filter_chain.empty())
        {
            arguments.push_back(L"-vf");
            arguments.push_back(filter_chain);
        }

        arguments.push_back(L"-c:a");
        arguments.push_back(L"copy");
        arguments.push_back(L"-sn");
        arguments.push_back(L"-dn");

        if (LowerExtension(output_path) == L".mp4")
        {
            arguments.push_back(L"-movflags");
            arguments.push_back(L"+faststart");
        }

        arguments.push_back(output_path.wstring());
        return arguments;
    }

    [[nodiscard]] TranscodeAttempt TranscodeMediaFile(
        const std::filesystem::path& ffmpeg_path,
        const std::filesystem::path& input_path,
        const std::filesystem::path& final_output_path,
        std::wstring_view encoder,
        const pptxvp::AppConfig& config,
        const LiveProgressCallback& progress_callback)
    {
        const std::filesystem::path temporary_output_path = MakeTemporaryOutputPath(final_output_path);
        const std::vector<std::wstring> arguments =
            BuildTranscodeArguments(input_path, temporary_output_path, encoder, config);
        FfmpegProgressParser parser;

        const pptxvp::helper::ProcessResult result = pptxvp::helper::RunProcessStreaming(
            ffmpeg_path,
            arguments,
            [&](std::string_view chunk)
            {
                parser.Consume(chunk, progress_callback);
            },
            input_path.parent_path());

        if (result.exit_code != 0)
        {
            std::error_code error_code;
            std::filesystem::remove(temporary_output_path, error_code);

            return TranscodeAttempt{
                .succeeded = false,
                .final_path = final_output_path,
                .detail = ShortFailureDetail(result.output),
            };
        }

        if (final_output_path == input_path)
        {
            std::filesystem::remove(input_path);
            std::filesystem::rename(temporary_output_path, input_path);
        }
        else
        {
            std::filesystem::rename(temporary_output_path, final_output_path);
            std::filesystem::remove(input_path);
        }

        return TranscodeAttempt{
            .succeeded = true,
            .final_path = final_output_path,
            .detail = L"处理成功。",
        };
    }
}

namespace pptxvp
{
    MediaProcessSummary ProcessMedia(
        const std::filesystem::path& extracted_root,
        const std::filesystem::path& ffmpeg_path,
        const AppConfig& config)
    {
        MediaProcessSummary summary;
        const std::filesystem::path media_directory = extracted_root / "ppt" / "media";
        const std::vector<std::filesystem::path> candidates = FindCandidateMediaFiles(media_directory);
        const auto processing_started_at = std::chrono::steady_clock::now();

        for (std::size_t index = 0; index < candidates.size(); ++index)
        {
            const std::filesystem::path& media_path = candidates[index];
            MediaActionResult action;
            action.source_path = media_path;
            action.output_path = media_path;

            helper::WriteProgressLine(BuildProgressLine(
                index + 1,
                candidates.size(),
                media_path.filename().wstring(),
                L"分析中",
                std::nullopt,
                std::nullopt,
                L"",
                std::chrono::steady_clock::now() - processing_started_at));

            const ProbeInfo probe = ProbeMedia(ffmpeg_path, media_path);

            if (!probe.has_video)
            {
                action.status = MediaActionStatus::Skipped;
                action.message = L"ffmpeg 未检测到视频流，已跳过。";
                helper::WriteProgressLine(BuildProgressLine(
                    index + 1,
                    candidates.size(),
                    media_path.filename().wstring(),
                    L"已跳过",
                    std::nullopt,
                    std::nullopt,
                    L"",
                    std::chrono::steady_clock::now() - processing_started_at));
                ++summary.skipped_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            const std::optional<std::wstring> encoder = ResolveEffectiveEncoder(config, probe);

            if (!encoder.has_value())
            {
                action.status = MediaActionStatus::Skipped;
                action.message = config.HasVideoChanges()
                                     ? L"无法将原始视频编码映射到可用编码器，已跳过。"
                                     : L"当前未配置任何视频处理参数，因此保持原样。";
                helper::WriteProgressLine(BuildProgressLine(
                    index + 1,
                    candidates.size(),
                    media_path.filename().wstring(),
                    L"已跳过",
                    std::nullopt,
                    probe.duration_seconds,
                    L"",
                    std::chrono::steady_clock::now() - processing_started_at));
                ++summary.skipped_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            const CodecFamily encoder_family = CodecFamilyFromEncoder(*encoder);
            std::filesystem::path preferred_output_path = media_path;
            bool should_retry_with_mp4 = false;

            if (encoder_family != CodecFamily::Unknown)
            {
                if (!ContainerSupportsCodec(media_path, encoder_family))
                {
                    preferred_output_path = MakeConvertedTargetPath(media_path);
                }
            }
            else
            {
                should_retry_with_mp4 = LowerExtension(media_path) != L".mp4";
            }

            helper::WriteProgressLine(BuildProgressLine(
                index + 1,
                candidates.size(),
                media_path.filename().wstring(),
                L"",
                std::nullopt,
                probe.duration_seconds,
                L"",
                std::chrono::steady_clock::now() - processing_started_at));

            TranscodeAttempt attempt = TranscodeMediaFile(
                ffmpeg_path,
                media_path,
                preferred_output_path,
                *encoder,
                config,
                [&](const LiveProgressState& state)
                {
                    helper::WriteProgressLine(BuildProgressLine(
                        index + 1,
                        candidates.size(),
                        media_path.filename().wstring(),
                        L"",
                        state.current_seconds,
                        probe.duration_seconds,
                        state.speed,
                        std::chrono::steady_clock::now() - processing_started_at));
                });

            if (!attempt.succeeded && preferred_output_path == media_path && should_retry_with_mp4)
            {
                preferred_output_path = MakeConvertedTargetPath(media_path);
                attempt = TranscodeMediaFile(
                    ffmpeg_path,
                    media_path,
                    preferred_output_path,
                    *encoder,
                    config,
                    [&](const LiveProgressState& state)
                    {
                        helper::WriteProgressLine(BuildProgressLine(
                            index + 1,
                            candidates.size(),
                            media_path.filename().wstring(),
                            L"",
                            state.current_seconds,
                            probe.duration_seconds,
                            state.speed,
                            std::chrono::steady_clock::now() - processing_started_at));
                    });
            }

            if (!attempt.succeeded)
            {
                action.status = MediaActionStatus::Failed;
                action.message = attempt.detail;
                helper::WriteProgressLine(BuildProgressLine(
                    index + 1,
                    candidates.size(),
                    media_path.filename().wstring(),
                    L"失败",
                    std::nullopt,
                    probe.duration_seconds,
                    L"",
                    std::chrono::steady_clock::now() - processing_started_at));
                ++summary.failed_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            action.status = MediaActionStatus::Processed;
            action.output_path = attempt.final_path;
            action.message = attempt.final_path == media_path
                                 ? L"视频内容已在原位置更新。"
                                 : L"视频容器已改为 MP4，并同步更新了 PPTX 引用。";
            helper::WriteProgressLine(BuildProgressLine(
                index + 1,
                candidates.size(),
                media_path.filename().wstring(),
                L"",
                probe.duration_seconds,
                probe.duration_seconds,
                L"",
                std::chrono::steady_clock::now() - processing_started_at));

            if (attempt.final_path != media_path)
            {
                summary.renames.push_back(MediaRename{
                    .original_relative_path = std::filesystem::relative(media_path, extracted_root),
                    .updated_relative_path = std::filesystem::relative(attempt.final_path, extracted_root),
                });
            }

            ++summary.processed_count;
            summary.items.push_back(std::move(action));
        }

        helper::FinishProgressLine();
        return summary;
    }
}
