namespace PptxVideoProcessing.WinUI;

internal sealed class ProcessingOptionsConfig
{
    public string Encoder { get; set; } = string.Empty;

    public string HardwareAcceleration { get; set; } = "auto";

    public string Preset { get; set; } = string.Empty;

    public string FrameRate { get; set; } = string.Empty;

    public string Resolution { get; set; } = string.Empty;

    public int VolumePercent { get; set; } = 100;

    public bool Mute { get; set; }
}
