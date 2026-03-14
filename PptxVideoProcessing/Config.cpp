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
}

namespace pptxvp
{
    AppConfig LoadConfig(const std::filesystem::path& config_path)
    {
        if (!std::filesystem::exists(config_path))
        {
            helper::WriteTextFileUtf8(config_path, "{}");
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

        return config;
    }
}
