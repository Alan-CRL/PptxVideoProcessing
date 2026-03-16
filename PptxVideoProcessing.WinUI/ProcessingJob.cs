using System.ComponentModel;
using System.Runtime.CompilerServices;
using Microsoft.UI.Xaml;

namespace PptxVideoProcessing.WinUI;

public sealed class ProcessingJob : INotifyPropertyChanged
{
    private string _status = "待处理";
    private string _detail = "等待开始处理。";
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
                OnPropertyChanged(nameof(IsFinished));
                OnPropertyChanged(nameof(PrimaryActionText));
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

    public bool IsPending => Status == "待处理";

    public bool IsProcessing => Status == "处理中";

    public bool IsFinished => Status is "成功" or "失败" or "跳过";

    public bool HasDetails => !string.IsNullOrWhiteSpace(DetailSummary);

    public Visibility DetailsButtonVisibility => HasDetails ? Visibility.Visible : Visibility.Collapsed;

    public Visibility ExpandedDetailsVisibility => HasDetails && IsDetailsExpanded ? Visibility.Visible : Visibility.Collapsed;

    public string DetailsActionText => IsDetailsExpanded ? "收起详情" : "查看详情";

    public string PrimaryActionText => IsProcessing ? "取消" : "删除";

    public event PropertyChangedEventHandler? PropertyChanged;

    public void MarkPending(string detail = "等待开始处理。")
    {
        ClearResultSummary();
        Status = "待处理";
        Detail = detail;
    }

    public void MarkProcessing(string detail)
    {
        ClearResultSummary();
        Status = "处理中";
        Detail = detail;
    }

    public void MarkSucceeded(string detail)
    {
        Status = "成功";
        Detail = detail;
    }

    public void MarkFailed(string detail)
    {
        Status = "失败";
        Detail = detail;
    }

    public void MarkSkipped(string detail)
    {
        Status = "跳过";
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
