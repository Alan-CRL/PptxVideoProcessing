module;

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

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
        bool has_audio{};
        std::wstring codec_name;
        std::optional<double> duration_seconds;
        std::optional<double> frame_rate;
        std::optional<int> width;
        std::optional<int> height;
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

    [[nodiscard]] std::optional<double> CalculateFilePercent(
        const std::optional<double>& current_seconds,
        const std::optional<double>& total_seconds)
    {
        if (!current_seconds.has_value() || !total_seconds.has_value() || *total_seconds <= 0.0)
        {
            return std::nullopt;
        }

        return std::clamp(*current_seconds / *total_seconds, 0.0, 1.0) * 100.0;
    }

    void ReportMediaProgress(
        const pptxvp::MediaProgressCallback& progress_callback,
        std::size_t current_index,
        std::size_t total_count,
        const std::filesystem::path& media_path,
        std::wstring_view note,
        const std::optional<double>& current_seconds,
        const std::optional<double>& total_seconds,
        std::wstring_view speed,
        std::wstring_view acceleration_backend,
        std::chrono::steady_clock::duration elapsed)
    {
        if (progress_callback)
        {
            progress_callback(pptxvp::MediaProgressInfo{
                .media_path = media_path,
                .current_index = current_index,
                .total_count = total_count,
                .current_seconds = current_seconds,
                .total_seconds = total_seconds,
                .file_percent = CalculateFilePercent(current_seconds, total_seconds),
                .note = std::wstring(note),
                .speed = std::wstring(speed),
                .acceleration_backend = std::wstring(acceleration_backend),
            });
            return;
        }

        pptxvp::helper::WriteProgressLine(BuildProgressLine(
            current_index,
            total_count,
            media_path.filename().wstring(),
            note,
            current_seconds,
            total_seconds,
            speed,
            elapsed));
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

    enum class ResolvedAccelerationBackend
    {
        Software,
        Nvidia,
        IntelQsv,
        AmdAmf,
        MediaFoundation,
    };

    using EncoderOperationalCache = std::unordered_map<std::wstring, bool>;

    [[nodiscard]] bool IsCanonicalSoftwareEncoder(std::wstring_view encoder)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(encoder);
        return lowered == L"libx264" || lowered == L"libx265" || lowered == L"libsvtav1" || lowered == L"mpeg4";
    }

    [[nodiscard]] ResolvedAccelerationBackend AccelerationBackendFromEncoder(std::wstring_view encoder)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(encoder);

        if (lowered.contains(L"_nvenc"))
        {
            return ResolvedAccelerationBackend::Nvidia;
        }

        if (lowered.contains(L"_qsv"))
        {
            return ResolvedAccelerationBackend::IntelQsv;
        }

        if (lowered.contains(L"_amf"))
        {
            return ResolvedAccelerationBackend::AmdAmf;
        }

        if (lowered.contains(L"_mf"))
        {
            return ResolvedAccelerationBackend::MediaFoundation;
        }

        return ResolvedAccelerationBackend::Software;
    }

    [[nodiscard]] std::wstring DescribeAccelerationBackend(ResolvedAccelerationBackend backend)
    {
        switch (backend)
        {
        case ResolvedAccelerationBackend::Nvidia:
            return L"NVIDIA NVENC";
        case ResolvedAccelerationBackend::IntelQsv:
            return L"Intel QSV";
        case ResolvedAccelerationBackend::AmdAmf:
            return L"AMD AMF";
        case ResolvedAccelerationBackend::MediaFoundation:
            return L"系统原生 (Media Foundation)";
        default:
            return L"纯软件";
        }
    }

    [[nodiscard]] std::unordered_set<std::wstring> QueryAvailableEncoders(const std::filesystem::path& ffmpeg_path)
    {
        std::unordered_set<std::wstring> encoders;
        const pptxvp::helper::ProcessResult result =
            pptxvp::helper::RunProcess(ffmpeg_path, {L"-hide_banner", L"-encoders"});

        std::istringstream stream{result.output};
        std::string line;

        while (std::getline(stream, line))
        {
            std::istringstream line_stream(line);
            std::string flags;
            std::string encoder_name;

            if (!(line_stream >> flags >> encoder_name))
            {
                continue;
            }

            if (flags.size() >= 6)
            {
                encoders.insert(pptxvp::helper::ToLowerAscii(pptxvp::helper::Utf8ToWide(encoder_name)));
            }
        }

        return encoders;
    }

    [[nodiscard]] bool SupportsEncoder(
        const std::unordered_set<std::wstring>& encoders,
        std::wstring_view encoder_name)
    {
        return encoders.contains(pptxvp::helper::ToLowerAscii(encoder_name));
    }

    void AppendBackendIfMissing(
        std::vector<ResolvedAccelerationBackend>& order,
        ResolvedAccelerationBackend backend)
    {
        if (std::ranges::find(order, backend) == order.end())
        {
            order.push_back(backend);
        }
    }

    [[nodiscard]] std::wstring DescribeRequestedAcceleration(const pptxvp::AppConfig& config)
    {
        if (config.encoder.has_value())
        {
            const ResolvedAccelerationBackend explicit_backend = AccelerationBackendFromEncoder(*config.encoder);

            if (explicit_backend != ResolvedAccelerationBackend::Software)
            {
                return DescribeAccelerationBackend(explicit_backend);
            }
        }

        if (!config.hardware_acceleration.has_value() ||
            *config.hardware_acceleration == pptxvp::HardwareAcceleration::Auto)
        {
            return L"自动识别中";
        }

        switch (*config.hardware_acceleration)
        {
        case pptxvp::HardwareAcceleration::None:
            return DescribeAccelerationBackend(ResolvedAccelerationBackend::Software);
        case pptxvp::HardwareAcceleration::Nvidia:
            return DescribeAccelerationBackend(ResolvedAccelerationBackend::Nvidia);
        case pptxvp::HardwareAcceleration::IntelQsv:
            return DescribeAccelerationBackend(ResolvedAccelerationBackend::IntelQsv);
        case pptxvp::HardwareAcceleration::AmdAmf:
            return DescribeAccelerationBackend(ResolvedAccelerationBackend::AmdAmf);
        case pptxvp::HardwareAcceleration::MediaFoundation:
            return DescribeAccelerationBackend(ResolvedAccelerationBackend::MediaFoundation);
        default:
            return DescribeAccelerationBackend(ResolvedAccelerationBackend::Software);
        }
    }

    [[nodiscard]] std::vector<ResolvedAccelerationBackend> BuildBackendOrder(
        const pptxvp::AppConfig& config,
        std::wstring_view base_encoder)
    {
        std::vector<ResolvedAccelerationBackend> order;
        const ResolvedAccelerationBackend explicit_backend = AccelerationBackendFromEncoder(base_encoder);

        if (explicit_backend != ResolvedAccelerationBackend::Software)
        {
            AppendBackendIfMissing(order, explicit_backend);
        }

        if (!config.hardware_acceleration.has_value() ||
            *config.hardware_acceleration == pptxvp::HardwareAcceleration::Auto)
        {
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::Nvidia);
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::IntelQsv);
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::AmdAmf);
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::MediaFoundation);
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::Software);
            return order;
        }

        switch (*config.hardware_acceleration)
        {
        case pptxvp::HardwareAcceleration::None:
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::Software);
            return order;
        case pptxvp::HardwareAcceleration::Nvidia:
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::Nvidia);
            break;
        case pptxvp::HardwareAcceleration::IntelQsv:
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::IntelQsv);
            break;
        case pptxvp::HardwareAcceleration::AmdAmf:
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::AmdAmf);
            break;
        case pptxvp::HardwareAcceleration::MediaFoundation:
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::MediaFoundation);
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::Software);
            return order;
        default:
            AppendBackendIfMissing(order, ResolvedAccelerationBackend::Software);
            return order;
        }

        AppendBackendIfMissing(order, ResolvedAccelerationBackend::Nvidia);
        AppendBackendIfMissing(order, ResolvedAccelerationBackend::IntelQsv);
        AppendBackendIfMissing(order, ResolvedAccelerationBackend::AmdAmf);
        AppendBackendIfMissing(order, ResolvedAccelerationBackend::MediaFoundation);
        AppendBackendIfMissing(order, ResolvedAccelerationBackend::Software);
        return order;
    }

    [[nodiscard]] std::wstring ResolveHardwareAcceleratedEncoder(
        std::wstring_view encoder,
        ResolvedAccelerationBackend backend)
    {
        if (!IsCanonicalSoftwareEncoder(encoder))
        {
            return std::wstring(encoder);
        }

        const CodecFamily family = CodecFamilyFromEncoder(encoder);

        switch (backend)
        {
        case ResolvedAccelerationBackend::Nvidia:
            switch (family)
            {
            case CodecFamily::H264:
                return L"h264_nvenc";
            case CodecFamily::H265:
                return L"hevc_nvenc";
            case CodecFamily::Av1:
                return L"av1_nvenc";
            default:
                return std::wstring(encoder);
            }
        case ResolvedAccelerationBackend::IntelQsv:
            switch (family)
            {
            case CodecFamily::H264:
                return L"h264_qsv";
            case CodecFamily::H265:
                return L"hevc_qsv";
            case CodecFamily::Av1:
                return L"av1_qsv";
            default:
                return std::wstring(encoder);
            }
        case ResolvedAccelerationBackend::AmdAmf:
            switch (family)
            {
            case CodecFamily::H264:
                return L"h264_amf";
            case CodecFamily::H265:
                return L"hevc_amf";
            case CodecFamily::Av1:
                return L"av1_amf";
            default:
                return std::wstring(encoder);
            }
        case ResolvedAccelerationBackend::MediaFoundation:
            switch (family)
            {
            case CodecFamily::H264:
                return L"h264_mf";
            case CodecFamily::H265:
                return L"hevc_mf";
            case CodecFamily::Av1:
                return L"av1_mf";
            default:
                return std::wstring(encoder);
            }
        default:
            return std::wstring(encoder);
        }
    }

    void TryAddEncoderCandidate(
        std::vector<std::wstring>& candidates,
        const std::filesystem::path& ffmpeg_path,
        const std::unordered_set<std::wstring>& available_encoders,
        std::wstring_view encoder,
        EncoderOperationalCache& encoder_operational_cache)
    {
        if (encoder.empty())
        {
            return;
        }

        const std::wstring normalized = pptxvp::helper::ToLowerAscii(encoder);

        if (!SupportsEncoder(available_encoders, normalized))
        {
            return;
        }

        if (AccelerationBackendFromEncoder(normalized) != ResolvedAccelerationBackend::Software)
        {
            auto [it, inserted] = encoder_operational_cache.try_emplace(normalized, false);

            if (inserted)
            {
                const pptxvp::helper::ProcessResult probe_result = pptxvp::helper::RunProcess(
                    ffmpeg_path,
                    {
                        L"-hide_banner",
                        L"-loglevel",
                        L"error",
                        L"-f",
                        L"lavfi",
                        L"-i",
                        // NVENC rejects very small frames such as 128x72 on some GPUs/drivers.
                        L"color=size=256x144:rate=1:color=black",
                        L"-frames:v",
                        L"1",
                        L"-an",
                        L"-c:v",
                        normalized,
                        L"-f",
                        L"null",
                        L"-",
                    },
                    ffmpeg_path.parent_path());
                it->second = probe_result.exit_code == 0;
            }

            if (!it->second)
            {
                return;
            }
        }

        if (std::ranges::find(candidates, normalized) == candidates.end())
        {
            candidates.push_back(normalized);
        }
    }

    [[nodiscard]] std::vector<std::wstring> BuildEncoderCandidates(
        const std::filesystem::path& ffmpeg_path,
        std::wstring_view base_encoder,
        const pptxvp::AppConfig& config,
        const std::unordered_set<std::wstring>& available_encoders,
        EncoderOperationalCache& encoder_operational_cache)
    {
        std::vector<std::wstring> candidates;
        const std::wstring normalized_base_encoder = pptxvp::helper::ToLowerAscii(base_encoder);
        const ResolvedAccelerationBackend explicit_encoder_backend = AccelerationBackendFromEncoder(normalized_base_encoder);
        const CodecFamily family = CodecFamilyFromEncoder(normalized_base_encoder);
        const std::optional<std::wstring> software_encoder = EncoderForCodecFamily(family);
        const bool can_remap_to_other_backends =
            explicit_encoder_backend != ResolvedAccelerationBackend::Software ||
            IsCanonicalSoftwareEncoder(normalized_base_encoder);

        for (const ResolvedAccelerationBackend backend : BuildBackendOrder(config, normalized_base_encoder))
        {
            std::wstring candidate_encoder;

            if (backend == explicit_encoder_backend && explicit_encoder_backend != ResolvedAccelerationBackend::Software)
            {
                candidate_encoder = normalized_base_encoder;
            }
            else if (backend == ResolvedAccelerationBackend::Software)
            {
                if (explicit_encoder_backend != ResolvedAccelerationBackend::Software)
                {
                    if (!software_encoder.has_value())
                    {
                        continue;
                    }

                    candidate_encoder = *software_encoder;
                }
                else
                {
                    candidate_encoder = normalized_base_encoder;
                }
            }
            else
            {
                if (!can_remap_to_other_backends || !software_encoder.has_value())
                {
                    continue;
                }

                candidate_encoder = ResolveHardwareAcceleratedEncoder(*software_encoder, backend);
            }

            TryAddEncoderCandidate(
                candidates,
                ffmpeg_path,
                available_encoders,
                candidate_encoder,
                encoder_operational_cache);
        }

        if (candidates.empty())
        {
            candidates.push_back(normalized_base_encoder);
        }

        return candidates;
    }
    [[nodiscard]] std::wstring JoinFailureDetails(const std::vector<std::wstring>& details)
    {
        if (details.empty())
        {
            return L"ffmpeg 执行失败。";
        }

        std::wstring joined;

        for (std::size_t index = 0; index < details.size(); ++index)
        {
            if (index != 0)
            {
                joined.append(L"；");
            }

            joined.append(details[index]);
        }

        return joined;
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
        static const std::regex audio_regex(R"(Audio:\s*([A-Za-z0-9_]+))");
        static const std::regex duration_regex(R"(Duration:\s*([0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?))");
        static const std::regex dimension_regex(R"((\d{2,5})x(\d{2,5}))");
        static const std::regex fps_regex(R"(([0-9]+(?:\.[0-9]+)?)\s*fps)");
        std::smatch match;
        std::optional<double> duration_seconds;
        std::optional<double> frame_rate;
        std::optional<int> width;
        std::optional<int> height;

        if (std::regex_search(result.output, match, duration_regex))
        {
            duration_seconds = ParseTimestampSeconds(match[1].str());
        }

        if (std::regex_search(result.output, match, dimension_regex))
        {
            width = std::stoi(match[1].str());
            height = std::stoi(match[2].str());
        }

        if (std::regex_search(result.output, match, fps_regex))
        {
            frame_rate = std::stod(match[1].str());
        }

        if (std::regex_search(result.output, match, video_regex))
        {
            return ProbeInfo{
                .has_video = true,
                .has_audio = std::regex_search(result.output, audio_regex),
                .codec_name = pptxvp::helper::Utf8ToWide(match[1].str()),
                .duration_seconds = duration_seconds,
                .frame_rate = frame_rate,
                .width = width,
                .height = height,
            };
        }

        return ProbeInfo{};
    }

    [[nodiscard]] bool MatchesTargetVideoConditions(const pptxvp::AppConfig& config, const ProbeInfo& probe)
    {
        if (!probe.has_video)
        {
            return false;
        }

        if (config.encoder.has_value())
        {
            const CodecFamily current_family = CodecFamilyFromDecoder(probe.codec_name);
            const CodecFamily target_family = CodecFamilyFromEncoder(*config.encoder);

            if (target_family != CodecFamily::Unknown && current_family != target_family)
            {
                return false;
            }
        }

        if (config.frame_rate.has_value())
        {
            if (!probe.frame_rate.has_value() || std::abs(*probe.frame_rate - static_cast<double>(*config.frame_rate)) > 0.05)
            {
                return false;
            }
        }

        if (config.resolution_height.has_value())
        {
            if (!probe.height.has_value() || *probe.height > *config.resolution_height)
            {
                return false;
            }
        }

        return true;
    }

    [[nodiscard]] std::optional<std::wstring> ResolveVideoEncoder(const pptxvp::AppConfig& config, const ProbeInfo& probe)
    {
        if (config.encoder.has_value())
        {
            return config.encoder;
        }

        if (!config.HasMediaChanges())
        {
            return std::nullopt;
        }

        return EncoderForCodecFamily(CodecFamilyFromDecoder(probe.codec_name));
    }

    [[nodiscard]] std::wstring BuildVideoFilterChain(const pptxvp::AppConfig& config)
    {
        std::vector<std::wstring> filters;

        if (config.frame_rate.has_value())
        {
            filters.push_back(L"fps=" + std::to_wstring(*config.frame_rate));
        }

        if (config.resolution_height.has_value())
        {
            const std::wstring max_height = std::to_wstring(*config.resolution_height);
            filters.push_back(
                L"scale='if(gt(ih," + max_height + L"),trunc(iw*" + max_height + L"/ih/2)*2,iw)':'if(gt(ih," + max_height + L")," + max_height + L",ih)'");
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

    [[nodiscard]] std::wstring BuildAudioVolumeExpression(const pptxvp::AppConfig& config)
    {
        if (config.mute)
        {
            return L"0";
        }

        const int volume_percent = config.volume_percent.value_or(100);
        std::wostringstream stream;
        stream << std::fixed << std::setprecision(2) << static_cast<double>(volume_percent) / 100.0;
        return stream.str();
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
        const std::optional<std::wstring>& video_encoder,
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
        };

        if (video_encoder.has_value())
        {
            arguments.push_back(L"-c:v");
            arguments.push_back(*video_encoder);
        }
        else
        {
            arguments.push_back(L"-c:v");
            arguments.push_back(L"copy");
        }

        if (video_encoder.has_value() && config.preset.has_value())
        {
            arguments.push_back(L"-preset");
            arguments.push_back(*config.preset);
        }

        const std::wstring filter_chain = BuildVideoFilterChain(config);

        if (!filter_chain.empty())
        {
            arguments.push_back(L"-vf");
            arguments.push_back(filter_chain);
        }

        if (config.HasAudioChanges())
        {
            arguments.push_back(L"-c:a");
            arguments.push_back(L"aac");
            arguments.push_back(L"-b:a");
            arguments.push_back(L"192k");
            arguments.push_back(L"-af");
            arguments.push_back(L"volume=" + BuildAudioVolumeExpression(config));
        }
        else
        {
            arguments.push_back(L"-c:a");
            arguments.push_back(L"copy");
        }

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
        const std::optional<std::wstring>& video_encoder,
        const pptxvp::AppConfig& config,
        const LiveProgressCallback& progress_callback)
    {
        const std::filesystem::path temporary_output_path = MakeTemporaryOutputPath(final_output_path);
        const std::vector<std::wstring> arguments =
            BuildTranscodeArguments(input_path, temporary_output_path, video_encoder, config);
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
        const AppConfig& config,
        const MediaProgressCallback& progress_callback)
    {
        MediaProcessSummary summary;
        const std::filesystem::path media_directory = extracted_root / "ppt" / "media";
        const std::vector<std::filesystem::path> candidates = FindCandidateMediaFiles(media_directory);
        const auto processing_started_at = std::chrono::steady_clock::now();
        const std::unordered_set<std::wstring> available_encoders = QueryAvailableEncoders(ffmpeg_path);
        EncoderOperationalCache encoder_operational_cache;
        summary.acceleration_backend = DescribeRequestedAcceleration(config);

        for (std::size_t index = 0; index < candidates.size(); ++index)
        {
            const std::filesystem::path& media_path = candidates[index];
            MediaActionResult action;
            action.source_path = media_path;
            action.output_path = media_path;

            ReportMediaProgress(
                progress_callback,
                index + 1,
                candidates.size(),
                media_path,
                L"分析中",
                std::nullopt,
                std::nullopt,
                L"",
                summary.acceleration_backend,
                std::chrono::steady_clock::now() - processing_started_at);

            const ProbeInfo probe = ProbeMedia(ffmpeg_path, media_path);

            if (!probe.has_video)
            {
                action.status = MediaActionStatus::Skipped;
                action.message = L"ffmpeg 未检测到视频流，已跳过。";
                ReportMediaProgress(
                    progress_callback,
                    index + 1,
                    candidates.size(),
                    media_path,
                    L"已跳过",
                    std::nullopt,
                    std::nullopt,
                    L"",
                    summary.acceleration_backend,
                    std::chrono::steady_clock::now() - processing_started_at);
                ++summary.skipped_count;
                ++summary.no_video_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            pptxvp::AppConfig effective_config = config;

            if (!probe.has_audio)
            {
                effective_config.volume_percent.reset();
                effective_config.mute = false;
            }

            if (config.HasAudioChanges() && !probe.has_audio && !config.HasVideoChanges())
            {
                action.status = MediaActionStatus::Skipped;
                action.message = L"视频不包含音频流，无法应用音量设置。";
                ReportMediaProgress(
                    progress_callback,
                    index + 1,
                    candidates.size(),
                    media_path,
                    L"已跳过",
                    probe.duration_seconds,
                    probe.duration_seconds,
                    L"",
                    summary.acceleration_backend,
                    std::chrono::steady_clock::now() - processing_started_at);
                ++summary.skipped_count;
                ++summary.already_satisfied_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            if (MatchesTargetVideoConditions(config, probe) && !effective_config.HasAudioChanges())
            {
                action.status = MediaActionStatus::Skipped;
                action.message = L"视频已满足目标条件，保持原样。";
                ReportMediaProgress(
                    progress_callback,
                    index + 1,
                    candidates.size(),
                    media_path,
                    L"已满足目标条件",
                    probe.duration_seconds,
                    probe.duration_seconds,
                    L"",
                    summary.acceleration_backend,
                    std::chrono::steady_clock::now() - processing_started_at);
                ++summary.skipped_count;
                ++summary.already_satisfied_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            const std::optional<std::wstring> encoder = ResolveVideoEncoder(effective_config, probe);

            if (!encoder.has_value())
            {
                action.status = MediaActionStatus::Skipped;
                action.message = effective_config.HasMediaChanges()
                                     ? L"无法将原始视频编码映射到可用编码器，已跳过。"
                                     : L"当前未配置任何处理参数，因此保持原样。";
                ReportMediaProgress(
                    progress_callback,
                    index + 1,
                    candidates.size(),
                    media_path,
                    L"已跳过",
                    std::nullopt,
                    probe.duration_seconds,
                    L"",
                    summary.acceleration_backend,
                    std::chrono::steady_clock::now() - processing_started_at);
                ++summary.skipped_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            const std::vector<std::wstring> encoder_candidates =
                BuildEncoderCandidates(ffmpeg_path, *encoder, effective_config, available_encoders, encoder_operational_cache);
            const std::wstring preferred_acceleration_label = encoder_candidates.empty()
                ? summary.acceleration_backend
                : DescribeAccelerationBackend(AccelerationBackendFromEncoder(encoder_candidates.front()));
            TranscodeAttempt attempt;
            std::wstring used_encoder;
            std::wstring acceleration_label = preferred_acceleration_label;
            std::wstring fallback_note;
            std::vector<std::wstring> failure_details;

            for (std::size_t candidate_index = 0; candidate_index < encoder_candidates.size(); ++candidate_index)
            {
                const std::wstring& candidate_encoder = encoder_candidates[candidate_index];
                const std::wstring candidate_acceleration_label =
                    DescribeAccelerationBackend(AccelerationBackendFromEncoder(candidate_encoder));
                const CodecFamily encoder_family = CodecFamilyFromEncoder(candidate_encoder);
                std::filesystem::path preferred_output_path = media_path;
                bool should_retry_with_mp4 =
                    effective_config.HasAudioChanges() && LowerExtension(media_path) != L".mp4";

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

                ReportMediaProgress(
                    progress_callback,
                    index + 1,
                    candidates.size(),
                    media_path,
                    candidate_index == 0 ? L"" : L"首选加速不可用，正在自动回退...",
                    std::nullopt,
                    probe.duration_seconds,
                    L"",
                    candidate_acceleration_label,
                    std::chrono::steady_clock::now() - processing_started_at);

                TranscodeAttempt candidate_attempt = TranscodeMediaFile(
                    ffmpeg_path,
                    media_path,
                    preferred_output_path,
                    candidate_encoder,
                    effective_config,
                    [&](const LiveProgressState& state)
                    {
                        ReportMediaProgress(
                            progress_callback,
                            index + 1,
                            candidates.size(),
                            media_path,
                            L"",
                            state.current_seconds,
                            probe.duration_seconds,
                            state.speed,
                            candidate_acceleration_label,
                            std::chrono::steady_clock::now() - processing_started_at);
                    });

                if (!candidate_attempt.succeeded && preferred_output_path == media_path && should_retry_with_mp4)
                {
                    preferred_output_path = MakeConvertedTargetPath(media_path);
                    candidate_attempt = TranscodeMediaFile(
                        ffmpeg_path,
                        media_path,
                        preferred_output_path,
                        candidate_encoder,
                        effective_config,
                        [&](const LiveProgressState& state)
                        {
                            ReportMediaProgress(
                                progress_callback,
                                index + 1,
                                candidates.size(),
                                media_path,
                                L"",
                                state.current_seconds,
                                probe.duration_seconds,
                                state.speed,
                                candidate_acceleration_label,
                                std::chrono::steady_clock::now() - processing_started_at);
                        });
                }

                if (candidate_attempt.succeeded)
                {
                    attempt = std::move(candidate_attempt);
                    used_encoder = candidate_encoder;
                    acceleration_label = candidate_acceleration_label;
                    break;
                }

                failure_details.push_back(candidate_acceleration_label + L"：" + candidate_attempt.detail);
            }

            if (used_encoder.empty())
            {
                action.status = MediaActionStatus::Failed;
                action.message = JoinFailureDetails(failure_details);
                ReportMediaProgress(
                    progress_callback,
                    index + 1,
                    candidates.size(),
                    media_path,
                    L"失败",
                    std::nullopt,
                    probe.duration_seconds,
                    L"",
                    acceleration_label,
                    std::chrono::steady_clock::now() - processing_started_at);
                ++summary.failed_count;
                summary.items.push_back(std::move(action));
                continue;
            }

            summary.acceleration_backend = acceleration_label;

            if (acceleration_label != preferred_acceleration_label)
            {
                fallback_note = L"首选加速不可用，已自动回退到 " + acceleration_label + L"。";
            }

            action.status = MediaActionStatus::Processed;
            action.output_path = attempt.final_path;
            action.message = attempt.final_path == media_path
                                 ? L"视频内容已在原位置更新。"
                                 : L"视频容器已改为 MP4。";

            if (!fallback_note.empty())
            {
                action.message.append(L" ");
                action.message.append(fallback_note);
            }

            ReportMediaProgress(
                progress_callback,
                index + 1,
                candidates.size(),
                media_path,
                L"",
                probe.duration_seconds,
                probe.duration_seconds,
                L"",
                acceleration_label,
                std::chrono::steady_clock::now() - processing_started_at);

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

        if (!progress_callback)
        {
            helper::FinishProgressLine();
        }

        return summary;
    }
}





















