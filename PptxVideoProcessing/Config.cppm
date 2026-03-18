export module PptxVideoProcessing.Config;

import std.compat;

export namespace pptxvp
{
    enum class HardwareAcceleration
    {
        Auto,
        None,
        Nvidia,
        IntelQsv,
        AmdAmf,
        MediaFoundation,
    };

    struct AppConfig
    {
        std::optional<std::wstring> encoder;
        std::optional<int> frame_rate;
        std::optional<int> resolution_height;
        std::optional<HardwareAcceleration> hardware_acceleration;
        std::optional<std::wstring> preset;
        std::optional<int> volume_percent;
        bool mute{};

        [[nodiscard]] bool HasVideoChanges() const noexcept
        {
            return encoder.has_value() || frame_rate.has_value() || resolution_height.has_value();
        }

        [[nodiscard]] bool HasAudioChanges() const noexcept
        {
            return mute || (volume_percent.has_value() && *volume_percent != 100);
        }

        [[nodiscard]] bool HasMediaChanges() const noexcept
        {
            return HasVideoChanges() || HasAudioChanges();
        }
    };

    [[nodiscard]] AppConfig LoadConfig(const std::filesystem::path& config_path);
}
