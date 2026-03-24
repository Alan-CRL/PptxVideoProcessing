module;

#include <nlohmann/json.hpp>

module PptxVideoProcessing.Config;

import std.compat;

import PptxVideoProcessing.Helper.FileSystem;
import PptxVideoProcessing.Helper.Utf;

namespace
{
    using json = nlohmann::json;

    [[nodiscard]] std::runtime_error MakeConfigError(std::wstring_view message)
    {
        return std::runtime_error(pptxvp::helper::WideToUtf8(message));
    }

    [[nodiscard]] std::string_view TrimAsciiWhitespace(std::string_view text)
    {
        const auto is_not_space = [](unsigned char character)
        {
            return !std::isspace(character);
        };

        const auto first = std::find_if(text.begin(), text.end(), is_not_space);
        const auto last = std::find_if(text.rbegin(), text.rend(), is_not_space).base();

        if (first >= last)
        {
            return {};
        }

        return text.substr(
            static_cast<std::size_t>(std::distance(text.begin(), first)),
            static_cast<std::size_t>(std::distance(first, last)));
    }

    [[nodiscard]] std::wstring NormalizeEncoder(std::wstring_view encoder)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(encoder);

        if (lowered.empty())
        {
            throw MakeConfigError(L"config.json 中的 encoder 不能为空字符串。");
        }

        if (lowered == L"h264")
        {
            return L"libx264";
        }

        if (lowered == L"h265")
        {
            return L"libx265";
        }

        if (lowered == L"av1")
        {
            return L"libsvtav1";
        }

        if (lowered == L"mpeg4")
        {
            return L"mpeg4";
        }

        return lowered;
    }

    [[nodiscard]] int ParseResolution(std::wstring_view resolution)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(resolution);

        if (lowered == L"360p")
        {
            return 360;
        }

        if (lowered == L"480p")
        {
            return 480;
        }

        if (lowered == L"720p")
        {
            return 720;
        }

        if (lowered == L"1080p")
        {
            return 1080;
        }

        if (lowered == L"2160p")
        {
            return 2160;
        }

        throw MakeConfigError(
            L"config.json 中的 resolution 只能是 360p、480p、720p、1080p 或 2160p。");
    }

    [[nodiscard]] pptxvp::HardwareAcceleration ParseHardwareAcceleration(std::wstring_view value)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(value);

        if (lowered.empty())
        {
            throw MakeConfigError(L"config.json 中的 hardwareAcceleration 不能为空字符串。");
        }

        if (lowered == L"auto")
        {
            return pptxvp::HardwareAcceleration::Auto;
        }

        if (lowered == L"none")
        {
            return pptxvp::HardwareAcceleration::None;
        }

        if (lowered == L"nvidia" || lowered == L"nvenc")
        {
            return pptxvp::HardwareAcceleration::Nvidia;
        }

        if (lowered == L"intel" || lowered == L"qsv" || lowered == L"intelqsv")
        {
            return pptxvp::HardwareAcceleration::IntelQsv;
        }

        if (lowered == L"amd" || lowered == L"amf")
        {
            return pptxvp::HardwareAcceleration::AmdAmf;
        }

        if (lowered == L"windows" || lowered == L"mf" || lowered == L"mediafoundation")
        {
            return pptxvp::HardwareAcceleration::MediaFoundation;
        }

        throw MakeConfigError(L"config.json 中的 hardwareAcceleration 只能是 auto、none、nvidia、intel、amd 或 windows。");
    }

    [[nodiscard]] std::wstring NormalizePreset(std::wstring_view preset)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(preset);

        if (lowered.empty())
        {
            throw MakeConfigError(L"config.json 中的 preset 不能为空字符串。");
        }

        return lowered;
    }

    [[nodiscard]] pptxvp::PresetLevel ParsePresetLevel(std::wstring_view value)
    {
        const std::wstring lowered = pptxvp::helper::ToLowerAscii(value);

        if (lowered.empty())
        {
            throw MakeConfigError(L"config.json 中的 presetLevel 不能为空字符串。");
        }

        if (lowered == L"low")
        {
            return pptxvp::PresetLevel::Low;
        }

        if (lowered == L"medium")
        {
            return pptxvp::PresetLevel::Medium;
        }

        if (lowered == L"high")
        {
            return pptxvp::PresetLevel::High;
        }

        throw MakeConfigError(L"config.json 中的 presetLevel 只能是 low、medium 或 high。");
    }
}

namespace pptxvp
{
    AppConfig LoadConfig(const std::filesystem::path& config_path)
    {
        if (!std::filesystem::exists(config_path))
        {
            helper::WriteTextFileUtf8(config_path, "{\"hardwareAcceleration\":\"auto\"}\n");
        }

        const std::string content = helper::ReadTextFileUtf8(config_path);
        const std::string_view trimmed_content = TrimAsciiWhitespace(content);

        json document;

        try
        {
            document = trimmed_content.empty() ? json::object() : json::parse(trimmed_content);
        }
        catch (const json::exception& exception)
        {
            throw std::runtime_error("解析 config.json 失败。 " + std::string(exception.what()));
        }

        if (!document.is_object())
        {
            throw MakeConfigError(L"config.json 的根节点必须是 JSON 对象。");
        }

        AppConfig config;

        if (document.contains("encoder"))
        {
            if (!document["encoder"].is_string())
            {
                throw MakeConfigError(L"config.json 中的 encoder 必须是字符串。");
            }

            config.encoder = NormalizeEncoder(helper::Utf8ToWide(document["encoder"].get<std::string>()));
        }

        if (document.contains("frameRate"))
        {
            if (!document["frameRate"].is_number_integer())
            {
                throw MakeConfigError(L"config.json 中的 frameRate 必须是整数。");
            }

            const int frame_rate = document["frameRate"].get<int>();

            if (frame_rate <= 0)
            {
                throw MakeConfigError(L"config.json 中的 frameRate 必须大于 0。");
            }

            config.frame_rate = frame_rate;
        }

        if (document.contains("resolution"))
        {
            if (!document["resolution"].is_string())
            {
                throw MakeConfigError(L"config.json 中的 resolution 必须是字符串。");
            }

            config.resolution_height = ParseResolution(helper::Utf8ToWide(document["resolution"].get<std::string>()));
        }

        if (document.contains("hardwareAcceleration"))
        {
            if (!document["hardwareAcceleration"].is_string())
            {
                throw MakeConfigError(L"config.json 中的 hardwareAcceleration 必须是字符串。");
            }

            config.hardware_acceleration =
                ParseHardwareAcceleration(helper::Utf8ToWide(document["hardwareAcceleration"].get<std::string>()));
        }
        else
        {
            config.hardware_acceleration = HardwareAcceleration::Auto;
        }

        if (document.contains("preset"))
        {
            if (!document["preset"].is_string())
            {
                throw MakeConfigError(L"config.json 中的 preset 必须是字符串。");
            }

            config.preset = NormalizePreset(helper::Utf8ToWide(document["preset"].get<std::string>()));
        }

        if (document.contains("presetLevel"))
        {
            if (!document["presetLevel"].is_string())
            {
                throw MakeConfigError(L"config.json 中的 presetLevel 必须是字符串。");
            }

            config.preset_level =
                ParsePresetLevel(helper::Utf8ToWide(document["presetLevel"].get<std::string>()));
        }

        if (document.contains("volumePercent"))
        {
            if (!document["volumePercent"].is_number_integer())
            {
                throw MakeConfigError(L"config.json 中的 volumePercent 必须是整数。");
            }

            const int volume_percent = document["volumePercent"].get<int>();

            if (volume_percent < 0 || volume_percent > 300)
            {
                throw MakeConfigError(L"config.json 中的 volumePercent 必须在 0 到 300 之间。");
            }

            config.volume_percent = volume_percent;
        }

        if (document.contains("mute"))
        {
            if (!document["mute"].is_boolean())
            {
                throw MakeConfigError(L"config.json 中的 mute 必须是布尔值。");
            }

            config.mute = document["mute"].get<bool>();
        }

        return config;
    }
}


