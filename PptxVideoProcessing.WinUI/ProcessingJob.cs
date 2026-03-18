using System.ComponentModel;
using System.Runtime.CompilerServices;
using Microsoft.UI.Xaml;

namespace PptxVideoProcessing.WinUI;

public sealed class ProcessingJob : INotifyPropertyChanged
{
    private const string PendingStatus = "未开始";
    private const string ProcessingStatus = "处理中";
    private const string PausedStatus = "已暂停";
    private const string StoppedStatus = "已停止";
    private const string FailedStatus = "失败";
    private const string SkippedStatus = "跳过";
    private const string SuccessStatus = "成功";

    private string _status = PendingStatus;
    private string _detail = "等待开始。";
    private string _detailSummary = string.Empty;
    private bool _isDetailsExpanded;

    public ProcessingJob(string inputPath)
    {
        InputPath = inputPath;
        FileName = Path.GetFileName(inputPath);
    }

    public string InputPath { get; }

    public string FileName { get; }

    public string Status
    {
        get => _status;
        private set
        {
            if (SetProperty(ref _status, value))
            {
                OnPropertyChanged(nameof(IsPending));
                OnPropertyChanged(nameof(IsProcessing));
                OnPropertyChanged(nameof(IsPaused));
                OnPropertyChanged(nameof(IsStopped));
                OnPropertyChanged(nameof(IsFinished));
                OnPropertyChanged(nameof(PrimaryActionText));
                OnPropertyChanged(nameof(PrimaryActionVisibility));
                OnPropertyChanged(nameof(SecondaryActionText));
                OnPropertyChanged(nameof(SecondaryActionVisibility));
            }
        }
    }

    public string Detail
    {
        get => _detail;
        private set => SetProperty(ref _detail, value);
    }

    public string DetailSummary
    {
        get => _detailSummary;
        private set
        {
            if (SetProperty(ref _detailSummary, value))
            {
                OnPropertyChanged(nameof(HasDetails));
                OnPropertyChanged(nameof(DetailsButtonVisibility));
                OnPropertyChanged(nameof(ExpandedDetailsVisibility));
                OnPropertyChanged(nameof(DetailsActionText));

                if (string.IsNullOrWhiteSpace(value))
                {
                    IsDetailsExpanded = false;
                }
            }
        }
    }

    public bool IsDetailsExpanded
    {
        get => _isDetailsExpanded;
        set
        {
            if (SetProperty(ref _isDetailsExpanded, value))
            {
                OnPropertyChanged(nameof(ExpandedDetailsVisibility));
                OnPropertyChanged(nameof(DetailsActionText));
            }
        }
    }

    public bool IsPending => Status == PendingStatus;

    public bool IsProcessing => Status == ProcessingStatus;

    public bool IsPaused => Status == PausedStatus;

    public bool IsStopped => Status == StoppedStatus;

    public bool IsFinished => Status is SuccessStatus or FailedStatus or SkippedStatus or StoppedStatus;

    public bool HasDetails => !string.IsNullOrWhiteSpace(DetailSummary);

    public Visibility PrimaryActionVisibility => GetPrimaryActionText() is null
        ? Visibility.Collapsed
        : Visibility.Visible;

    public Visibility SecondaryActionVisibility => GetSecondaryActionText() is null
        ? Visibility.Collapsed
        : Visibility.Visible;

    public Visibility DetailsButtonVisibility => HasDetails ? Visibility.Visible : Visibility.Collapsed;

    public Visibility ExpandedDetailsVisibility => HasDetails && IsDetailsExpanded ? Visibility.Visible : Visibility.Collapsed;

    public string DetailsActionText => IsDetailsExpanded ? "收起详情" : "查看详情";

    public string PrimaryActionText => GetPrimaryActionText() ?? string.Empty;

    public string SecondaryActionText => GetSecondaryActionText() ?? string.Empty;

    public event PropertyChangedEventHandler? PropertyChanged;

    public void MarkPending(string detail = "等待开始。")
    {
        ClearResultSummary();
        Status = PendingStatus;
        Detail = detail;
    }

    public void MarkProcessing(string detail)
    {
        ClearResultSummary();
        Status = ProcessingStatus;
        Detail = detail;
    }

    public void MarkPaused(string detail)
    {
        Status = PausedStatus;
        Detail = detail;
    }

    public void MarkStopped(string detail)
    {
        Status = StoppedStatus;
        Detail = detail;
    }

    public void MarkSucceeded(string detail)
    {
        Status = SuccessStatus;
        Detail = detail;
    }

    public void MarkFailed(string detail)
    {
        Status = FailedStatus;
        Detail = detail;
    }

    public void MarkSkipped(string detail)
    {
        Status = SkippedStatus;
        Detail = detail;
    }

    public void SetResultSummary(string summary, bool expand = false)
    {
        DetailSummary = summary;
        IsDetailsExpanded = expand && HasDetails;
    }

    public void ToggleDetails()
    {
        if (!HasDetails)
        {
            return;
        }

        IsDetailsExpanded = !IsDetailsExpanded;
    }

    public void ClearResultSummary()
    {
        DetailSummary = string.Empty;
        IsDetailsExpanded = false;
    }

    private string? GetPrimaryActionText()
    {
        return Status switch
        {
            PendingStatus => "开始",
            ProcessingStatus => "暂停",
            PausedStatus => "继续",
            StoppedStatus or FailedStatus or SkippedStatus or SuccessStatus => "重试",
            _ => null,
        };
    }

    private string? GetSecondaryActionText()
    {
        return Status switch
        {
            PendingStatus or StoppedStatus or FailedStatus or SkippedStatus or SuccessStatus => "删除",
            ProcessingStatus or PausedStatus => "停止",
            _ => null,
        };
    }

    private bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(storage, value))
        {
            return false;
        }

        storage = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
