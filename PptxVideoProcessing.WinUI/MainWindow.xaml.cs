using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Graphics;
using WinRT.Interop;

namespace PptxVideoProcessing.WinUI;

public sealed partial class MainWindow : Window
{
    private const string PendingStatus = "待处理";
    private const string ProcessingStatus = "处理中";
    private const string FailedStatus = "失败";
    private const string SkippedStatus = "跳过";
    private const string SuccessStatus = "成功";
    private const string SystemNativeAcceleration = "系统原生 (Media Foundation)";
    private const string AutoAcceleration = "自动识别中";
    private const string SoftwareAcceleration = "纯软件";
    private const int OpenFileDialogBufferSize = 65536;
    private const int DefaultWindowWidth = 1240;
    private const int DefaultWindowHeight = 1000;
    private const int MinimumWindowWidth = 980;
    private const int MinimumWindowHeight = 900;
    private const int OFN_FILEMUSTEXIST = 0x00001000;
    private const int OFN_PATHMUSTEXIST = 0x00000800;
    private const int OFN_ALLOWMULTISELECT = 0x00000200;
    private const int OFN_EXPLORER = 0x00080000;
    private const int OFN_HIDEREADONLY = 0x00000004;
    private const int OFN_NOCHANGEDIR = 0x00000008;
    private const int GWL_WNDPROC = -4;
    private const uint WM_GETMINMAXINFO = 0x0024;
    private const uint WM_NCDESTROY = 0x0082;

    private readonly string _workerPath = Path.Combine(App.RuntimeAssetsDirectory, "PptxVideoProcessing.Worker.exe");
    private readonly string _ffmpegPath = Path.Combine(App.RuntimeAssetsDirectory, "ffmpeg.exe");
    private readonly string _configPath = Path.Combine(App.RuntimeAssetsDirectory, "config.json");
    private readonly List<ProcessingJob> _activeBatchJobs = [];
    private AppWindow? _appWindow;
    private Process? _currentWorkerProcess;
    private ProcessingJob? _currentJob;
    private Stopwatch? _currentJobStopwatch;
    private ProcessingJob? _cancellationTargetJob;
    private bool _allowWindowClose;
    private bool _closeConfirmationPending;
    private bool _configHasVideoChanges;
    private bool _isProcessing;
    private string _configuredAccelerationLabel = SystemNativeAcceleration;
    private string _activeAccelerationLabel = SystemNativeAcceleration;
    private bool _stopQueueAfterCurrentJob;
    private bool _queueClearedDuringProcessing;
    private double? _currentFilePercent;
    private double _currentJobProgressFraction;
    private nint _windowHandle;
    private nint _originalWindowProcedure;
    private WindowProcedure? _windowProcedureDelegate;

    public MainWindow()
    {
        InitializeComponent();
        ConfigureWindow();
        Closed += OnWindowClosed;
        ResetIdleState();
        UpdateCommandStates();
        App.WriteDiagnosticLog("主窗口", "MainWindow 已初始化。");

        if (!File.Exists(_ffmpegPath))
        {
            string message = $"程序目录未找到 ffmpeg.exe。请将 ffmpeg.exe 放到 {_ffmpegPath} 后再开始处理。";
            App.WriteDiagnosticLog("主窗口", message);
            ShowInfoBar(message, InfoBarSeverity.Warning);
        }
    }

    public ObservableCollection<ProcessingJob> QueueItems { get; } = new();

    private void ConfigureWindow()
    {
        _windowHandle = WindowNative.GetWindowHandle(this);
        var windowId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(_windowHandle);
        _appWindow = AppWindow.GetFromWindowId(windowId);
        _appWindow.Closing += OnAppWindowClosing;
        OverlappedPresenter? presenter = _appWindow.Presenter as OverlappedPresenter;

        if (presenter is null)
        {
            _appWindow.SetPresenter(AppWindowPresenterKind.Overlapped);
            presenter = _appWindow.Presenter as OverlappedPresenter;
        }

        if (presenter is not null)
        {
            presenter.IsResizable = true;
            presenter.IsMaximizable = true;
            presenter.IsMinimizable = true;
        }

        _appWindow.Resize(new SizeInt32(DefaultWindowWidth, DefaultWindowHeight));
        InstallMinimumWindowSizeHook();
    }

    private void SelectPptxButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            App.WriteDiagnosticLog("选择 PPTX", "用户点击了选择按钮。");
            IReadOnlyList<string> selectedPaths = PickPptxFiles();
            App.WriteDiagnosticLog("选择 PPTX", $"文件选择已返回，共 {selectedPaths.Count} 个文件。");

            if (selectedPaths.Count == 0)
            {
                App.WriteDiagnosticLog("选择 PPTX", "用户取消了选择或没有返回任何文件。");
                return;
            }

            AddPptxPathsToQueue(selectedPaths, "选择 PPTX");
        }
        catch (Exception exception)
        {
            HandleUiActionError("选择 PPTX 文件", exception);
        }
    }

    private void QueueListView_DragOver(object sender, DragEventArgs e)
    {
        if (e.DataView.Contains(StandardDataFormats.StorageItems) || e.DataView.Contains(StandardDataFormats.Text))
        {
            e.AcceptedOperation = DataPackageOperation.Copy;
            e.DragUIOverride.Caption = "拖入一个或多个 PPTX 文件";
            e.DragUIOverride.IsCaptionVisible = true;
            return;
        }

        e.AcceptedOperation = DataPackageOperation.None;
    }

    private async void QueueListView_Drop(object sender, DragEventArgs e)
    {
        var deferral = e.GetDeferral();

        try
        {
            IReadOnlyList<string> droppedPaths = await GetDroppedPptxPathsAsync(e.DataView);
            App.WriteDiagnosticLog("拖拽导入", $"拖拽解析完成，共 {droppedPaths.Count} 个 PPTX 文件。");

            if (droppedPaths.Count == 0)
            {
                ShowInfoBar("仅支持拖入一个或多个 .pptx 文件。", InfoBarSeverity.Informational);
                return;
            }

            AddPptxPathsToQueue(droppedPaths, "拖拽导入");
        }
        catch (Exception exception)
        {
            HandleUiActionError("拖拽导入 PPTX 文件", exception);
        }
        finally
        {
            deferral.Complete();
        }
    }

    private void AddPptxPathsToQueue(IReadOnlyList<string> selectedPaths, string sourceLabel)
    {
        int addedCount = AddJobsToQueue(selectedPaths);

        if (addedCount == 0)
        {
            ShowInfoBar("所选文件都已经在队列中了。", InfoBarSeverity.Informational);
        }
        else if (_isProcessing && !_stopQueueAfterCurrentJob)
        {
            RefreshBatchProgress();
            ShowInfoBar($"已加入 {addedCount} 个 PPTX 文件，当前批次会继续处理。", InfoBarSeverity.Success);
        }
        else if (_isProcessing)
        {
            ShowInfoBar($"已加入 {addedCount} 个 PPTX 文件。当前处理结束后，请重新点击开始处理。", InfoBarSeverity.Informational);
        }
        else
        {
            ResetIdleState();
            ShowInfoBar($"已加入 {addedCount} 个 PPTX 文件。", InfoBarSeverity.Success);
        }

        UpdateCommandStates();
        App.WriteDiagnosticLog(sourceLabel, $"队列更新完成，当前共有 {QueueItems.Count} 个文件。");
    }

    private static async Task<IReadOnlyList<string>> GetDroppedPptxPathsAsync(DataPackageView dataView)
    {
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (dataView.Contains(StandardDataFormats.StorageItems))
        {
            IReadOnlyList<IStorageItem> items = await dataView.GetStorageItemsAsync();

            foreach (IStorageItem item in items)
            {
                if (!item.IsOfType(StorageItemTypes.File) || string.IsNullOrWhiteSpace(item.Path))
                {
                    continue;
                }

                if (string.Equals(Path.GetExtension(item.Path), ".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    paths.Add(item.Path);
                }
            }
        }

        if (paths.Count == 0 && dataView.Contains(StandardDataFormats.Text))
        {
            string text = await dataView.GetTextAsync();
            string[] candidates = text.Split(["\r\n", "\n", "\r"], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            foreach (string rawCandidate in candidates)
            {
                string candidate = rawCandidate.Trim().Trim('"');

                if (!string.Equals(Path.GetExtension(candidate), ".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (File.Exists(candidate))
                {
                    paths.Add(candidate);
                }
            }
        }

        return paths.OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();
    }
    private async void StartProcessingButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await StartQueueProcessingAsync();
        }
        catch (Exception exception)
        {
            HandleUiActionError("启动处理", exception);
        }
    }

    private async void ConfigureOptionsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            await ShowProcessingOptionsDialogAsync();
        }
        catch (Exception exception)
        {
            HandleUiActionError("编辑处理选项", exception);
        }
    }

    private void RemoveJobButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ProcessingJob? job = GetJobFromSender(sender);

            if (job is null)
            {
                return;
            }

            if (_isProcessing && ReferenceEquals(job, _currentJob))
            {
                CancelCurrentJobAndRemove(job);
                return;
            }

            RemoveJobFromQueue(job, $"已移除 {job.FileName}。", InfoBarSeverity.Informational);
        }
        catch (Exception exception)
        {
            HandleUiActionError("删除任务", exception);
        }
    }

    private void ToggleDetailsButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ProcessingJob? job = GetJobFromSender(sender);

            if (job is null)
            {
                return;
            }

            job.ToggleDetails();
        }
        catch (Exception exception)
        {
            HandleUiActionError("切换任务详情", exception);
        }
    }
    private void ClearFinishedButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            List<ProcessingJob> jobsToRemove = QueueItems.Where(job => job.IsFinished).ToList();

            if (jobsToRemove.Count == 0)
            {
                ShowInfoBar("当前没有可清理的已完成结果。", InfoBarSeverity.Informational);
                return;
            }

            foreach (ProcessingJob job in jobsToRemove)
            {
                _activeBatchJobs.Remove(job);
                QueueItems.Remove(job);
            }

            App.WriteDiagnosticLog("队列操作", $"已清理 {jobsToRemove.Count} 个已完成结果。");
            HandleQueueMutationAfterRemoval();
            ShowInfoBar($"已清理 {jobsToRemove.Count} 个已完成/失败/跳过的任务。", InfoBarSeverity.Informational);
        }
        catch (Exception exception)
        {
            HandleUiActionError("清理结果", exception);
        }
    }

    private void ClearAllButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (QueueItems.Count == 0)
            {
                ShowInfoBar("当前队列已经是空的。", InfoBarSeverity.Informational);
                return;
            }

            int clearedCount = QueueItems.Count;
            App.WriteDiagnosticLog("队列操作", $"用户请求清空全部队列，共 {clearedCount} 项。");

            if (_isProcessing)
            {
                _queueClearedDuringProcessing = true;
                _stopQueueAfterCurrentJob = true;
                _currentJobProgressFraction = 0.0;
                _currentFilePercent = null;

                if (_currentJob is not null)
                {
                    _cancellationTargetJob = _currentJob;
                }

                QueueItems.Clear();
                _activeBatchJobs.Clear();
                ResetIdleState("当前状态：队列已清空", "当前 PPTX：无", "当前视频(已清空)");
                UpdateCommandStates();
                ShowInfoBar($"已清空 {clearedCount} 个队列项，正在终止当前处理。", InfoBarSeverity.Informational);
                TerminateCurrentWorker();
                return;
            }

            QueueItems.Clear();
            _activeBatchJobs.Clear();
            ResetIdleState("当前状态：队列已清空", "当前 PPTX：无", "当前视频(已清空)");
            UpdateCommandStates();
            ShowInfoBar($"已清空 {clearedCount} 个队列项。", InfoBarSeverity.Informational);
        }
        catch (Exception exception)
        {
            HandleUiActionError("清空队列", exception);
        }
    }

    private async void ExitButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!_isProcessing)
            {
                App.WriteDiagnosticLog("退出软件", "当前空闲，直接关闭窗口。");
                _allowWindowClose = true;
                Close();
                return;
            }

            bool shouldExit = await ConfirmExitWhileProcessingAsync();
            App.WriteDiagnosticLog("退出软件", $"处理中退出确认结果：{shouldExit}");

            if (!shouldExit)
            {
                return;
            }

            _allowWindowClose = true;
            TerminateCurrentWorker();
            Close();
        }
        catch (Exception exception)
        {
            HandleUiActionError("退出软件", exception);
        }
    }

    private void OnAppWindowClosing(AppWindow sender, AppWindowClosingEventArgs args)
    {
        if (_allowWindowClose || !_isProcessing)
        {
            return;
        }

        args.Cancel = true;

        if (_closeConfirmationPending)
        {
            return;
        }

        _closeConfirmationPending = true;
        _ = ConfirmWindowCloseFromTitleBarAsync();
    }

    private async Task ConfirmWindowCloseFromTitleBarAsync()
    {
        try
        {
            bool shouldExit = await ConfirmExitWhileProcessingAsync();
            App.WriteDiagnosticLog("退出软件", $"标题栏关闭确认结果：{shouldExit}");

            if (!shouldExit)
            {
                return;
            }

            _allowWindowClose = true;
            TerminateCurrentWorker();
            Close();
        }
        finally
        {
            _closeConfirmationPending = false;
        }
    }
    private int AddJobsToQueue(IReadOnlyList<string> selectedPaths)
    {
        var existing = new HashSet<string>(
            QueueItems.Select(item => item.InputPath),
            StringComparer.OrdinalIgnoreCase);

        var addedJobs = new List<ProcessingJob>();

        foreach (string path in selectedPaths)
        {
            if (!existing.Add(path))
            {
                continue;
            }

            var job = new ProcessingJob(path);
            QueueItems.Add(job);
            addedJobs.Add(job);
        }

        if (_isProcessing && !_stopQueueAfterCurrentJob)
        {
            _activeBatchJobs.AddRange(addedJobs);
        }

        return addedJobs.Count;
    }

    private async Task StartQueueProcessingAsync()
    {
        if (_isProcessing)
        {
            App.WriteDiagnosticLog("开始处理", "用户点击开始处理，但当前已有任务在运行。");
            return;
        }

        List<ProcessingJob> jobsToRun = QueueItems.Where(IsReadyToStart).ToList();

        if (jobsToRun.Count == 0)
        {
            App.WriteDiagnosticLog("开始处理", "当前没有可处理的 PPTX 文件。");
            ShowInfoBar("当前没有可处理的 PPTX 文件。", InfoBarSeverity.Informational);
            return;
        }

        if (!TryValidateProcessingEnvironment(out string? blockingMessage, out string? advisoryMessage))
        {
            string message = blockingMessage ?? "处理环境校验失败。";
            App.WriteDiagnosticLog("开始处理", message);
            CurrentStatusText.Text = "当前状态：无法开始";
            CurrentDetailText.Text = "当前视频(无法开始)";
            CurrentFileProgressBar.IsIndeterminate = false;
            CurrentFileProgressBar.Value = 0;
            ShowInfoBar(message, InfoBarSeverity.Error);
            return;
        }

        App.WriteDiagnosticLog("开始处理", $"准备处理 {jobsToRun.Count} 个 PPTX 文件。");
        _isProcessing = true;
        _stopQueueAfterCurrentJob = false;
        _queueClearedDuringProcessing = false;
        _cancellationTargetJob = null;
        _currentJob = null;
        _currentWorkerProcess = null;
        _currentJobStopwatch = null;
        _currentFilePercent = null;
        foreach (ProcessingJob job in jobsToRun)
        {
            if (job.Status == FailedStatus)
            {
                job.MarkPending("等待重新处理。");
            }
        }

        _activeBatchJobs.Clear();
        _activeBatchJobs.AddRange(jobsToRun);
        StatusInfoBar.IsOpen = false;
        UpdateCommandStates();
        ResetProgressDisplay(_activeBatchJobs.Count);

        if (!string.IsNullOrWhiteSpace(advisoryMessage))
        {
            App.WriteDiagnosticLog("开始处理", advisoryMessage!);
            ShowInfoBar(advisoryMessage!, InfoBarSeverity.Warning);
        }

        try
        {
            while (!_stopQueueAfterCurrentJob)
            {
                ProcessingJob? nextJob = GetNextBatchJob();

                if (nextJob is null)
                {
                    break;
                }

                await RunWorkerForJobAsync(nextJob);
            }

            if (!_queueClearedDuringProcessing)
            {
                if (_activeBatchJobs.Count == 0)
                {
                    ResetIdleState();
                }
                else
                {
                    ApplyBatchCompletionState(_activeBatchJobs.ToList());
                }
            }
        }
        finally
        {
            _isProcessing = false;
            _currentWorkerProcess = null;
            _currentJob = null;
            _currentJobStopwatch = null;
            _currentFilePercent = null;
            _currentJobProgressFraction = 0.0;
            _cancellationTargetJob = null;
            _stopQueueAfterCurrentJob = false;
            _queueClearedDuringProcessing = false;
            _activeBatchJobs.Clear();
            UpdateCommandStates();
        }
    }

    private ProcessingJob? GetNextBatchJob()
    {
        foreach (ProcessingJob job in _activeBatchJobs)
        {
            if (IsPendingForCurrentBatch(job))
            {
                return job;
            }
        }

        return null;
    }

    private async Task RunWorkerForJobAsync(ProcessingJob job)
    {
        App.WriteDiagnosticLog("处理任务", $"开始处理：{job.InputPath}");
        _currentJob = job;
        _currentJobStopwatch = Stopwatch.StartNew();
        _currentFilePercent = null;
        _currentJobProgressFraction = 0.0;
        job.MarkProcessing("正在启动处理进程...");
        RefreshBatchProgress();
        UpdateCurrentProgress("正在启动处理进程...", job.FileName, "当前视频(准备中)", null, _configuredAccelerationLabel);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = _workerPath,
                WorkingDirectory = App.RuntimeAssetsDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            },
            EnableRaisingEvents = true,
        };

        process.StartInfo.ArgumentList.Add("--input");
        process.StartInfo.ArgumentList.Add(job.InputPath);
        process.StartInfo.ArgumentList.Add("--json-progress");
        process.StartInfo.ArgumentList.Add("--no-pause");

        App.WriteDiagnosticLog("处理任务", $"准备启动 worker：{_workerPath}");

        try
        {
            if (!process.Start())
            {
                HandleJobFailure(job, "无法启动 native worker。", showInfoBar: true);
                return;
            }

            App.WriteDiagnosticLog("处理任务", $"worker 已启动，PID={process.Id}");
            _currentWorkerProcess = process;
            Task<string> stderrTask = process.StandardError.ReadToEndAsync();

            bool completedEventReceived = false;
            string? outputLine;

            while ((outputLine = await process.StandardOutput.ReadLineAsync()) is not null)
            {
                if (string.IsNullOrWhiteSpace(outputLine))
                {
                    continue;
                }

                try
                {
                    using JsonDocument document = JsonDocument.Parse(outputLine);
                    completedEventReceived = HandleWorkerMessage(job, document.RootElement) || completedEventReceived;
                }
                catch (JsonException)
                {
                    HandleJobFailure(job, $"处理 {job.FileName} 时收到了无法解析的进度输出。", showInfoBar: true);
                    App.WriteDiagnosticLog("处理任务", $"收到无法解析的 JSON：{outputLine}");
                }
            }

            await process.WaitForExitAsync();
            string stderr = await stderrTask;
            App.WriteDiagnosticLog("处理任务", $"worker 已退出，ExitCode={process.ExitCode}");

            if (!string.IsNullOrWhiteSpace(stderr))
            {
                App.WriteDiagnosticLog("处理任务", $"worker 标准错误输出：{stderr.Trim()}");
            }

            bool canceledByUser = ReferenceEquals(_cancellationTargetJob, job);

            if (canceledByUser)
            {
                App.WriteDiagnosticLog("处理任务", $"任务已由用户取消：{job.InputPath}");
                _cancellationTargetJob = null;
                RefreshBatchProgress();
                return;
            }

            if (process.ExitCode == 0)
            {
                if (!completedEventReceived && !job.IsFinished)
                {
                    job.MarkSucceeded("处理完成");
                    job.SetResultSummary(BuildSuccessDetailSummary(
                        job.InputPath,
                        string.Empty,
                        GetCurrentJobElapsed(),
                        processedCount: 0,
                        skippedCount: 0,
                        failedCount: 0,
                        alreadySatisfiedCount: 0));
                    UpdateCurrentProgress("处理完成", job.FileName, "当前视频(已完成)", 100.0, _activeAccelerationLabel);
                }

                RefreshBatchProgress();
                return;
            }

            if (job.IsFinished)
            {
                RefreshBatchProgress();
                return;
            }

            HandleJobFailure(job, ExtractWorkerError(stderr), showInfoBar: true);
        }
        finally
        {
            _currentWorkerProcess = null;
            _currentJob = null;
            _currentJobStopwatch = null;
            _currentFilePercent = null;
            _currentJobProgressFraction = 0.0;
        }
    }

    private bool HandleWorkerMessage(ProcessingJob job, JsonElement root)
    {
        if (ReferenceEquals(_cancellationTargetJob, job))
        {
            return false;
        }

        string type = root.TryGetProperty("type", out JsonElement typeElement)
            ? typeElement.GetString() ?? string.Empty
            : string.Empty;

        switch (type)
        {
        case "progress":
            HandleProgressMessage(job, root);
            return false;
        case "completed":
            HandleCompletedMessage(job, root);
            return true;
        case "error":
            HandleErrorMessage(job, root);
            return false;
        default:
            HandleJobFailure(job, $"处理 {job.FileName} 时收到了未知事件类型。", showInfoBar: true);
            App.WriteDiagnosticLog("处理任务", $"收到未知事件类型：{type}");
            return false;
        }
    }

    private void HandleProgressMessage(ProcessingJob job, JsonElement root)
    {
        string message = ReadString(root, "message", "正在处理...");
        int currentIndex = ReadInt(root, "currentIndex", 0);
        int mediaTotal = ReadInt(root, "totalCount", 0);
        double? filePercent = ReadDouble(root, "filePercent");
        string accelerationBackend = ReadString(root, "accelerationBackend", string.Empty);

        string queueDetail = BuildQueueProgressText(currentIndex, mediaTotal);
        string videoDetail = BuildCurrentVideoText(currentIndex, mediaTotal, filePercent);
        _currentJobProgressFraction = CalculateCurrentJobProgressFraction(currentIndex, mediaTotal, filePercent);
        job.MarkProcessing(queueDetail);
        UpdateCurrentProgress(message, job.FileName, videoDetail, filePercent, accelerationBackend);
        RefreshBatchProgress();
    }

    private void HandleCompletedMessage(ProcessingJob job, JsonElement root)
    {
        string outputPath = ReadString(root, "outputPath", string.Empty);
        int processedCount = ReadInt(root, "processedCount", 0);
        int skippedCount = ReadInt(root, "skippedCount", 0);
        int failedCount = ReadInt(root, "failedCount", 0);
        int alreadySatisfiedCount = ReadInt(root, "alreadySatisfiedCount", 0);
        int noVideoCount = ReadInt(root, "noVideoCount", 0);
        int totalMedia = processedCount + skippedCount + failedCount;
        string accelerationBackend = ReadString(root, "accelerationBackend", string.Empty);
        bool hasOutput = !string.IsNullOrWhiteSpace(outputPath);

        App.WriteDiagnosticLog("处理任务", $"收到完成事件：输出={(hasOutput ? outputPath : "<无输出>")}，成功 {processedCount}，跳过 {skippedCount}，失败 {failedCount}。");

        if (!hasOutput)
        {
            string reason = BuildNoOutputReason(totalMedia, skippedCount, failedCount, alreadySatisfiedCount, noVideoCount);

            if (failedCount > 0)
            {
                job.MarkFailed("未生成输出");
                job.SetResultSummary(BuildFailureDetailSummary(job.InputPath, GetCurrentJobElapsed(), reason));
                UpdateCurrentProgress("处理结束（未生成输出）", job.FileName, "当前视频(未生成输出)", 100.0, accelerationBackend);
            }
            else
            {
                job.MarkSkipped("已跳过");
                job.SetResultSummary(BuildSkippedDetailSummary(job.InputPath, outputPath, GetCurrentJobElapsed(), reason));
                UpdateCurrentProgress("已跳过", job.FileName, "当前视频(未执行转码)", 100.0, accelerationBackend);
            }

            RefreshBatchProgress();
            return;
        }

        string queueDetail = totalMedia > 0 ? $"{totalMedia}/{totalMedia} 视频" : "已完成";
        job.MarkSucceeded(queueDetail);
        job.SetResultSummary(BuildSuccessDetailSummary(
            job.InputPath,
            outputPath,
            GetCurrentJobElapsed(),
            processedCount,
            skippedCount,
            failedCount,
            alreadySatisfiedCount));
        UpdateCurrentProgress("处理完成", job.FileName, "当前视频(已完成)", 100.0, accelerationBackend);
        RefreshBatchProgress();
    }

    private void HandleErrorMessage(ProcessingJob job, JsonElement root)
    {
        string message = ReadString(root, "message", "处理失败。");
        HandleJobFailure(job, message, showInfoBar: true);
        App.WriteDiagnosticLog("处理任务", $"worker 报告错误：{message}");
    }

    private void HandleJobFailure(ProcessingJob job, string message, bool showInfoBar)
    {
        job.MarkFailed("处理失败");
        job.SetResultSummary(BuildFailureDetailSummary(job.InputPath, GetCurrentJobElapsed(), message));
        UpdateCurrentProgress("处理失败", job.FileName, "当前视频(处理失败)", 100.0, _activeAccelerationLabel);
        RefreshBatchProgress();

        if (showInfoBar)
        {
            ShowInfoBar(message, InfoBarSeverity.Error);
        }
    }

    private IReadOnlyList<string> PickPptxFiles()
    {
        IntPtr buffer = Marshal.AllocCoTaskMem(OpenFileDialogBufferSize * sizeof(char));

        try
        {
            byte[] zeroBytes = new byte[OpenFileDialogBufferSize * sizeof(char)];
            Marshal.Copy(zeroBytes, 0, buffer, zeroBytes.Length);

            var dialog = new OPENFILENAME
            {
                lStructSize = Marshal.SizeOf<OPENFILENAME>(),
                hwndOwner = WindowNative.GetWindowHandle(this),
                lpstrFilter = "PowerPoint 演示文稿 (*.pptx)\0*.pptx\0\0",
                lpstrFile = buffer,
                nMaxFile = OpenFileDialogBufferSize,
                lpstrTitle = "选择一个或多个 PPTX 文件",
                lpstrDefExt = "pptx",
                Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_HIDEREADONLY | OFN_NOCHANGEDIR,
            };

            if (!GetOpenFileName(ref dialog))
            {
                int errorCode = CommDlgExtendedError();

                if (errorCode == 0)
                {
                    return Array.Empty<string>();
                }

                throw new InvalidOperationException($"系统文件对话框返回错误 0x{errorCode:X4}。{DescribeOpenFileDialogError(errorCode)}");
            }

            string rawSelection = Marshal.PtrToStringUni(buffer, OpenFileDialogBufferSize) ?? string.Empty;
            IReadOnlyList<string> results = ParseSelectedFiles(rawSelection);
            App.WriteDiagnosticLog("选择 PPTX", $"系统文件对话框解析完成，共得到 {results.Count} 个文件。");
            return results;
        }
        finally
        {
            Marshal.FreeCoTaskMem(buffer);
        }
    }

    private async Task<bool> ConfirmExitWhileProcessingAsync()
    {
        ContentDialog dialog = new()
        {
            Title = "确认退出",
            Content = "当前仍有 PPTX 正在处理。退出后会终止当前处理并结束剩余队列，是否继续？",
            PrimaryButtonText = "退出并终止处理",
            CloseButtonText = "继续处理",
            DefaultButton = ContentDialogButton.Close,
            XamlRoot = RootGrid.XamlRoot,
        };

        return await dialog.ShowAsync() == ContentDialogResult.Primary;
    }

    private void ResetProgressDisplay(int totalJobs)
    {
        CurrentStatusText.Text = "当前状态：准备开始";
        CurrentPptxText.Text = "当前 PPTX：尚未开始";
        CurrentDetailText.Text = "当前视频(准备中)";
        SetCurrentAcceleration(_configuredAccelerationLabel);
        CurrentFileProgressBar.IsIndeterminate = false;
        CurrentFileProgressBar.Value = 0;
        _currentFilePercent = null;
        _currentJobProgressFraction = 0.0;
        BatchProgressBar.Value = 0;
        BatchProgressText.Text = $"0 / {totalJobs} 已完成";
    }

    private void ResetIdleState(
        string status = "当前状态：空闲",
        string currentPptx = "当前 PPTX：尚未开始",
        string currentVideo = "当前视频(准备中)")
    {
        CurrentStatusText.Text = status;
        CurrentPptxText.Text = currentPptx;
        CurrentDetailText.Text = currentVideo;
        SetCurrentAcceleration(_configuredAccelerationLabel);
        CurrentFileProgressBar.IsIndeterminate = false;
        CurrentFileProgressBar.Value = 0;
        _currentFilePercent = null;
        _currentJobProgressFraction = 0.0;
        BatchProgressBar.Value = 0;

        int pendingCount = QueueItems.Count(job => IsReadyToStart(job));
        BatchProgressText.Text = pendingCount == 0 ? "0 / 0 已完成" : $"0 / {pendingCount} 已完成";
    }

    private void SetCurrentAcceleration(string accelerationLabel)
    {
        _activeAccelerationLabel = string.IsNullOrWhiteSpace(accelerationLabel)
            ? SystemNativeAcceleration
            : accelerationLabel;
        CurrentAccelerationText.Text = $"当前加速：{_activeAccelerationLabel}";
    }

    private static string DescribeConfiguredAcceleration(string? rawValue)
    {
        if (string.IsNullOrWhiteSpace(rawValue))
        {
            return SystemNativeAcceleration;
        }

        string normalized = rawValue.Trim().ToLowerInvariant();

        if (normalized.Contains("_nvenc") || normalized is "nvidia" or "nvenc")
        {
            return "NVIDIA NVENC";
        }

        if (normalized.Contains("_qsv") || normalized is "intel" or "qsv" or "intelqsv")
        {
            return "Intel QSV";
        }

        if (normalized.Contains("_amf") || normalized is "amd" or "amf")
        {
            return "AMD AMF";
        }

        if (normalized.Contains("_mf") || normalized is "windows" or "mf" or "mediafoundation")
        {
            return SystemNativeAcceleration;
        }

        return normalized switch
        {
            "auto" => AutoAcceleration,
            "none" => SoftwareAcceleration,
            _ => SystemNativeAcceleration,
        };
    }

    private void UpdateCurrentProgress(
        string status,
        string pptxName,
        string detail,
        double? filePercent,
        string? accelerationLabel = null)
    {
        CurrentStatusText.Text = $"当前状态：{status}";
        CurrentPptxText.Text = $"当前 PPTX：{pptxName}";
        CurrentDetailText.Text = detail;
        _currentFilePercent = filePercent;

        if (!string.IsNullOrWhiteSpace(accelerationLabel))
        {
            SetCurrentAcceleration(accelerationLabel!);
        }

        if (filePercent.HasValue)
        {
            CurrentFileProgressBar.IsIndeterminate = false;
            CurrentFileProgressBar.Value = Math.Clamp(filePercent.Value, 0.0, 100.0);
        }
        else
        {
            CurrentFileProgressBar.IsIndeterminate = true;
        }
    }

    private void RefreshBatchProgress()
    {
        int totalJobs = _activeBatchJobs.Count;

        if (totalJobs <= 0)
        {
            BatchProgressBar.Value = 0;
            BatchProgressText.Text = "0 / 0 已完成";
            return;
        }

        int completedJobs = _activeBatchJobs.Count(job => job.IsFinished);
        double currentFraction = 0.0;

        if (_currentJob is not null && _activeBatchJobs.Contains(_currentJob) && _currentJob.Status == ProcessingStatus)
        {
            currentFraction = Math.Clamp(_currentJobProgressFraction, 0.0, 1.0);
        }

        double overallPercent = Math.Min(completedJobs + currentFraction, totalJobs) / totalJobs * 100.0;
        BatchProgressBar.Value = overallPercent;
        BatchProgressText.Text = $"{completedJobs} / {totalJobs} 已完成";
    }

    private static double CalculateCurrentJobProgressFraction(int currentIndex, int totalCount, double? filePercent)
    {
        if (totalCount <= 0)
        {
            return 0.0;
        }

        int displayIndex = Math.Clamp(currentIndex, 1, totalCount);
        double currentMediaFraction = Math.Clamp((filePercent ?? 0.0) / 100.0, 0.0, 1.0);
        return ((displayIndex - 1) + currentMediaFraction) / totalCount;
    }
    private void UpdateCommandStates()
    {
        SelectButton.IsEnabled = true;
        OptionsButton.IsEnabled = !_isProcessing;
        StartButton.IsEnabled = !_isProcessing && QueueItems.Any(IsReadyToStart);
        ClearFinishedButton.IsEnabled = QueueItems.Any(job => job.IsFinished);
        ClearAllButton.IsEnabled = QueueItems.Count > 0;
    }

    private async Task ShowProcessingOptionsDialogAsync()
    {
        ProcessingOptionsConfig options = LoadProcessingOptionsConfig();
        TextBox encoderTextBox = new()
        {
            Text = options.Encoder,
            PlaceholderText = "例如 h264、h265、av1、libx264、h264_nvenc",
        };
        ComboBox accelerationComboBox = CreateOptionsComboBox(
        [
            ("默认：系统原生", string.Empty),
            ("自动识别", "auto"),
            ("纯软件", "none"),
            ("NVIDIA", "nvidia"),
            ("Intel", "intel"),
            ("AMD", "amd"),
            ("系统原生", "windows"),
        ], options.HardwareAcceleration);
        TextBox presetTextBox = new()
        {
            Text = options.Preset,
            PlaceholderText = "留空使用编码器默认预设，例如 medium、fast、p4、p5",
        };
        TextBox frameRateTextBox = new()
        {
            Text = options.FrameRate,
            PlaceholderText = "例如 24、30、60",
        };
        ComboBox resolutionComboBox = CreateOptionsComboBox(
        [
            ("保持原分辨率", string.Empty),
            ("360p", "360p"),
            ("480p", "480p"),
            ("720p", "720p"),
            ("1080p", "1080p"),
            ("2160p", "2160p"),
        ], options.Resolution);
        TextBlock validationTextBlock = new()
        {
            TextWrapping = TextWrapping.WrapWholeWords,
            Visibility = Visibility.Collapsed,
        };

        StackPanel panel = new()
        {
            Spacing = 12,
        };
        panel.Children.Add(new TextBlock
        {
            Text = "空白表示恢复为默认。保存后会写回程序目录中的 config.json。",
            TextWrapping = TextWrapping.WrapWholeWords,
            Opacity = 0.78,
        });
        panel.Children.Add(CreateOptionsEditorField("视频编码器", "常用可写 h264、h265、av1、mpeg4，也支持直接写 libx264、h264_nvenc 这类 ffmpeg 编码器名。", encoderTextBox));
        panel.Children.Add(CreateOptionsEditorField("硬件加速", "缺省时优先尝试系统原生加速；“自动识别”会根据本机显卡和 ffmpeg 支持自动选择。", accelerationComboBox));
        panel.Children.Add(CreateOptionsEditorField("编码预设", "用于平衡速度与压缩效率。留空时使用编码器默认值；常见如 medium、fast、p4、p5。", presetTextBox));
        panel.Children.Add(CreateOptionsEditorField("目标帧率", "只填正整数，例如 24、30、60。留空表示不修改帧率。", frameRateTextBox));
        panel.Children.Add(CreateOptionsEditorField("目标分辨率", "按目标高度缩放并保持宽高比。留空表示保持原分辨率。", resolutionComboBox));
        panel.Children.Add(validationTextBlock);

        ContentDialog dialog = new()
        {
            Title = "处理选项",
            PrimaryButtonText = "保存",
            CloseButtonText = "取消",
            DefaultButton = ContentDialogButton.Primary,
            XamlRoot = RootGrid.XamlRoot,
            Content = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 620,
                Content = panel,
            },
        };

        ProcessingOptionsConfig? updatedOptions = null;
        dialog.PrimaryButtonClick += (_, args) =>
        {
            if (!TryCollectProcessingOptions(
                    encoderTextBox,
                    accelerationComboBox,
                    presetTextBox,
                    frameRateTextBox,
                    resolutionComboBox,
                    out ProcessingOptionsConfig collectedOptions,
                    out string? validationMessage))
            {
                validationTextBlock.Text = validationMessage ?? "处理选项无效。";
                validationTextBlock.Visibility = Visibility.Visible;
                args.Cancel = true;
                return;
            }

            validationTextBlock.Visibility = Visibility.Collapsed;
            updatedOptions = collectedOptions;
        };

        if (await dialog.ShowAsync() != ContentDialogResult.Primary || updatedOptions is null)
        {
            return;
        }

        SaveProcessingOptionsConfig(updatedOptions);
        bool configValid = TryInspectConfig(out string? blockingMessage, out string? advisoryMessage);

        if (!configValid)
        {
            throw new InvalidOperationException(blockingMessage ?? "保存后的配置无效。");
        }

        if (!_isProcessing)
        {
            ResetIdleState();
        }
        else
        {
            SetCurrentAcceleration(_configuredAccelerationLabel);
        }

        string successMessage = "处理选项已保存。";

        if (!string.IsNullOrWhiteSpace(advisoryMessage))
        {
            successMessage = $"{successMessage}{advisoryMessage}";
            ShowInfoBar(successMessage, InfoBarSeverity.Warning);
        }
        else
        {
            ShowInfoBar(successMessage, InfoBarSeverity.Success);
        }
    }

    private ProcessingOptionsConfig LoadProcessingOptionsConfig()
    {
        var options = new ProcessingOptionsConfig();

        if (!File.Exists(_configPath))
        {
            return options;
        }

        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(_configPath));

        if (document.RootElement.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException("config.json 的根节点必须是 JSON 对象。");
        }

        JsonElement root = document.RootElement;
        options.Encoder = ReadString(root, "encoder", string.Empty);
        options.HardwareAcceleration = ReadString(root, "hardwareAcceleration", string.Empty);
        options.Preset = ReadString(root, "preset", string.Empty);
        options.FrameRate = root.TryGetProperty("frameRate", out JsonElement frameRate) ? frameRate.ToString() : string.Empty;
        options.Resolution = ReadString(root, "resolution", string.Empty);
        return options;
    }

    private void SaveProcessingOptionsConfig(ProcessingOptionsConfig options)
    {
        JsonObject root = new();

        if (File.Exists(_configPath))
        {
            JsonNode? node = JsonNode.Parse(File.ReadAllText(_configPath));

            if (node is JsonObject objectNode)
            {
                root = objectNode;
            }
        }

        SetOrRemoveJsonString(root, "encoder", options.Encoder);
        SetOrRemoveJsonString(root, "hardwareAcceleration", options.HardwareAcceleration);
        SetOrRemoveJsonString(root, "preset", options.Preset);
        SetOrRemoveJsonString(root, "resolution", options.Resolution);

        if (string.IsNullOrWhiteSpace(options.FrameRate))
        {
            root.Remove("frameRate");
        }
        else
        {
            root["frameRate"] = int.Parse(options.FrameRate);
        }

        string json = root.ToJsonString(new JsonSerializerOptions
        {
            WriteIndented = true,
        });
        File.WriteAllText(_configPath, json + Environment.NewLine, new UTF8Encoding(false));
        App.WriteDiagnosticLog("处理选项", $"config.json 已更新：{_configPath}");
    }

    private static bool TryCollectProcessingOptions(
        TextBox encoderTextBox,
        ComboBox accelerationComboBox,
        TextBox presetTextBox,
        TextBox frameRateTextBox,
        ComboBox resolutionComboBox,
        out ProcessingOptionsConfig options,
        out string? validationMessage)
    {
        options = new ProcessingOptionsConfig
        {
            Encoder = NormalizeOptionalText(encoderTextBox.Text),
            HardwareAcceleration = GetSelectedComboBoxValue(accelerationComboBox),
            Preset = NormalizeOptionalText(presetTextBox.Text),
            FrameRate = NormalizeOptionalText(frameRateTextBox.Text),
            Resolution = GetSelectedComboBoxValue(resolutionComboBox),
        };

        if (!string.IsNullOrWhiteSpace(options.FrameRate) && (!int.TryParse(options.FrameRate, out int frameRate) || frameRate <= 0))
        {
            validationMessage = "目标帧率必须是大于 0 的整数。";
            return false;
        }

        validationMessage = null;
        return true;
    }

    private static FrameworkElement CreateOptionsEditorField(string title, string description, Control editor)
    {
        StackPanel panel = new()
        {
            Spacing = 6,
        };
        panel.Children.Add(new TextBlock
        {
            Text = title,
        });
        panel.Children.Add(editor);
        panel.Children.Add(new TextBlock
        {
            Text = description,
            TextWrapping = TextWrapping.WrapWholeWords,
            Opacity = 0.74,
            Style = Application.Current.Resources["CaptionTextBlockStyle"] as Style,
        });
        return panel;
    }

    private static ComboBox CreateOptionsComboBox(IEnumerable<(string Label, string Value)> items, string selectedValue)
    {
        ComboBox comboBox = new();

        foreach ((string label, string value) in items)
        {
            ComboBoxItem item = new()
            {
                Content = label,
                Tag = value,
            };
            comboBox.Items.Add(item);

            if (string.Equals(value, selectedValue, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = item;
            }
        }

        if (comboBox.SelectedItem is null && comboBox.Items.Count > 0)
        {
            comboBox.SelectedIndex = 0;
        }

        return comboBox;
    }

    private static string GetSelectedComboBoxValue(ComboBox comboBox)
    {
        return comboBox.SelectedItem is ComboBoxItem { Tag: string value }
            ? value
            : string.Empty;
    }

    private static string NormalizeOptionalText(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
    }

    private static void SetOrRemoveJsonString(JsonObject root, string propertyName, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            root.Remove(propertyName);
            return;
        }

        root[propertyName] = value;
    }

    private void ShowInfoBar(string message, InfoBarSeverity severity)
    {
        StatusInfoBar.Message = message;
        StatusInfoBar.Severity = severity;
        StatusInfoBar.IsOpen = true;
    }

    private void HandleUiActionError(string actionName, Exception exception)
    {
        App.WriteDiagnosticLog("主窗口操作失败", $"{actionName} 发生异常：{exception.Message}");
        App.WriteCrashLog($"主窗口操作失败：{actionName}", exception);
        ShowInfoBar($"{actionName}失败：{exception.Message}", InfoBarSeverity.Error);
    }

    private void CancelCurrentJobAndRemove(ProcessingJob job)
    {
        App.WriteDiagnosticLog("队列操作", $"用户取消当前任务：{job.InputPath}");
        _cancellationTargetJob = job;
        _activeBatchJobs.Remove(job);
        QueueItems.Remove(job);
        CurrentStatusText.Text = "当前状态：正在取消";
        CurrentPptxText.Text = $"当前 PPTX：{job.FileName}";
        CurrentDetailText.Text = "当前视频(正在取消)";
        CurrentFileProgressBar.IsIndeterminate = true;
        _currentFilePercent = null;
        RefreshBatchProgress();
        UpdateCommandStates();
        ShowInfoBar($"已取消并移除 {job.FileName}。", InfoBarSeverity.Informational);
        TerminateCurrentWorker();
    }

    private void RemoveJobFromQueue(ProcessingJob job, string infoMessage, InfoBarSeverity severity)
    {
        bool removed = QueueItems.Remove(job);
        _activeBatchJobs.Remove(job);

        if (!removed)
        {
            return;
        }

        App.WriteDiagnosticLog("队列操作", $"已移除任务：{job.InputPath}");
        HandleQueueMutationAfterRemoval();
        ShowInfoBar(infoMessage, severity);
    }

    private void HandleQueueMutationAfterRemoval()
    {
        if (_isProcessing)
        {
            RefreshBatchProgress();
        }
        else
        {
            ResetIdleState();
        }

        UpdateCommandStates();
    }

    private ProcessingJob? GetJobFromSender(object sender)
    {
        if (sender is FrameworkElement element)
        {
            if (element.Tag is ProcessingJob taggedJob)
            {
                return taggedJob;
            }

            if (element.DataContext is ProcessingJob dataContextJob)
            {
                return dataContextJob;
            }
        }

        return null;
    }

    private void TerminateCurrentWorker()
    {
        try
        {
            if (_currentWorkerProcess is { HasExited: false })
            {
                App.WriteDiagnosticLog("处理任务", "正在终止当前 worker。");
                _currentWorkerProcess.Kill(true);
            }
        }
        catch
        {
            // Ignore teardown failures while the app is closing.
        }
    }

    private void OnWindowClosed(object sender, WindowEventArgs args)
    {
        App.WriteDiagnosticLog("主窗口", "窗口正在关闭。");

        if (_appWindow is not null)
        {
            _appWindow.Closing -= OnAppWindowClosing;
        }

        RestoreWindowProcedure();
        TerminateCurrentWorker();
    }

    private void InstallMinimumWindowSizeHook()
    {
        if (_windowHandle == 0 || _windowProcedureDelegate is not null)
        {
            return;
        }

        _windowProcedureDelegate = WindowProcedureCallback;
        _originalWindowProcedure = SetWindowLongPtr(
            _windowHandle,
            GWL_WNDPROC,
            Marshal.GetFunctionPointerForDelegate(_windowProcedureDelegate));
    }

    private void RestoreWindowProcedure()
    {
        if (_windowHandle == 0 || _originalWindowProcedure == 0)
        {
            return;
        }

        SetWindowLongPtr(_windowHandle, GWL_WNDPROC, _originalWindowProcedure);
        _originalWindowProcedure = 0;
        _windowProcedureDelegate = null;
    }

    private nint WindowProcedureCallback(nint hwnd, uint message, nuint wParam, nint lParam)
    {
        if (message == WM_GETMINMAXINFO)
        {
            MINMAXINFO minMaxInfo = Marshal.PtrToStructure<MINMAXINFO>(lParam);
            minMaxInfo.ptMinTrackSize.x = MinimumWindowWidth;
            minMaxInfo.ptMinTrackSize.y = MinimumWindowHeight;
            Marshal.StructureToPtr(minMaxInfo, lParam, false);
            return 0;
        }

        if (message == WM_NCDESTROY)
        {
            nint result = CallOriginalWindowProcedure(hwnd, message, wParam, lParam);
            RestoreWindowProcedure();
            return result;
        }

        return CallOriginalWindowProcedure(hwnd, message, wParam, lParam);
    }

    private nint CallOriginalWindowProcedure(nint hwnd, uint message, nuint wParam, nint lParam)
    {
        return _originalWindowProcedure != 0
            ? CallWindowProc(_originalWindowProcedure, hwnd, message, wParam, lParam)
            : DefWindowProc(hwnd, message, wParam, lParam);
    }

    private static bool IsPendingForCurrentBatch(ProcessingJob job)
    {
        return job.Status == PendingStatus;
    }

    private static bool IsReadyToStart(ProcessingJob job)
    {
        return job.Status is PendingStatus or FailedStatus;
    }

    private bool TryValidateProcessingEnvironment(out string? blockingMessage, out string? advisoryMessage)
    {
        blockingMessage = null;
        advisoryMessage = null;
        _configHasVideoChanges = false;
        _configuredAccelerationLabel = SystemNativeAcceleration;

        if (!File.Exists(_workerPath))
        {
            blockingMessage = $"未找到 native worker：{_workerPath}";
            return false;
        }

        if (!File.Exists(_ffmpegPath))
        {
            blockingMessage = $"未在程序目录找到 ffmpeg.exe。请将 ffmpeg.exe 放到 {_ffmpegPath} 后再重试。";
            return false;
        }

        return TryInspectConfig(out blockingMessage, out advisoryMessage);
    }

    private bool TryInspectConfig(out string? blockingMessage, out string? advisoryMessage)
    {
        blockingMessage = null;
        advisoryMessage = null;
        _configHasVideoChanges = false;
        _configuredAccelerationLabel = SystemNativeAcceleration;

        if (!File.Exists(_configPath))
        {
            advisoryMessage = "未找到 config.json，程序会自动创建空配置；如果未设置 encoder、frameRate 或 resolution，本次不会生成新的输出文件。";
            return true;
        }

        try
        {
            using JsonDocument document = JsonDocument.Parse(File.ReadAllText(_configPath));

            if (document.RootElement.ValueKind != JsonValueKind.Object)
            {
                blockingMessage = "config.json 的根节点必须是 JSON 对象。";
                return false;
            }

            JsonElement root = document.RootElement;
            bool hasVideoChanges = false;
            string? encoderName = null;

            if (root.TryGetProperty("encoder", out JsonElement encoder))
            {
                if (encoder.ValueKind != JsonValueKind.String)
                {
                    blockingMessage = "config.json 中的 encoder 必须是字符串。";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(encoder.GetString()))
                {
                    blockingMessage = "config.json 中的 encoder 不能为空字符串。";
                    return false;
                }

                encoderName = encoder.GetString()!.Trim();
                hasVideoChanges = true;
            }

            if (root.TryGetProperty("frameRate", out JsonElement frameRate))
            {
                if (frameRate.ValueKind != JsonValueKind.Number || !frameRate.TryGetInt32(out int frameRateValue))
                {
                    blockingMessage = "config.json 中的 frameRate 必须是整数。";
                    return false;
                }

                if (frameRateValue <= 0)
                {
                    blockingMessage = "config.json 中的 frameRate 必须大于 0。";
                    return false;
                }

                hasVideoChanges = true;
            }

            if (root.TryGetProperty("resolution", out JsonElement resolution))
            {
                if (resolution.ValueKind != JsonValueKind.String)
                {
                    blockingMessage = "config.json 中的 resolution 必须是字符串。";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(resolution.GetString()))
                {
                    blockingMessage = "config.json 中的 resolution 不能为空字符串。";
                    return false;
                }

                hasVideoChanges = true;
            }

            if (root.TryGetProperty("hardwareAcceleration", out JsonElement hardwareAcceleration))
            {
                if (hardwareAcceleration.ValueKind != JsonValueKind.String)
                {
                    blockingMessage = "config.json 中的 hardwareAcceleration 必须是字符串。";
                    return false;
                }

                string? rawAcceleration = hardwareAcceleration.GetString();

                if (string.IsNullOrWhiteSpace(rawAcceleration))
                {
                    blockingMessage = "config.json 中的 hardwareAcceleration 不能为空字符串。";
                    return false;
                }

                string normalized = rawAcceleration.Trim().ToLowerInvariant();

                if (normalized is not ("auto" or "none" or "nvidia" or "nvenc" or "intel" or "qsv" or "intelqsv" or "amd" or "amf" or "windows" or "mf" or "mediafoundation"))
                {
                    blockingMessage = "config.json 中的 hardwareAcceleration 只能是 auto、none、nvidia、intel、amd 或 windows。";
                    return false;
                }

                _configuredAccelerationLabel = DescribeConfiguredAcceleration(normalized);
            }
            else if (!string.IsNullOrWhiteSpace(encoderName))
            {
                _configuredAccelerationLabel = DescribeConfiguredAcceleration(encoderName);
            }

            if (root.TryGetProperty("preset", out JsonElement preset))
            {
                if (preset.ValueKind != JsonValueKind.String)
                {
                    blockingMessage = "config.json 中的 preset 必须是字符串。";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(preset.GetString()))
                {
                    blockingMessage = "config.json 中的 preset 不能为空字符串。";
                    return false;
                }
            }

            _configHasVideoChanges = hasVideoChanges;

            if (!hasVideoChanges)
            {
                advisoryMessage = "当前 config.json 没有设置 encoder、frameRate 或 resolution；仅设置 hardwareAcceleration 或 preset 不会触发转码，本次不会生成新的输出文件。";
            }

            return true;
        }
        catch (JsonException exception)
        {
            blockingMessage = $"解析 config.json 失败。{exception.Message}";
            return false;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            blockingMessage = $"读取 config.json 失败。{exception.Message}";
            return false;
        }
    }

    private void ApplyBatchCompletionState(IReadOnlyList<ProcessingJob> jobsToRun)
    {
        int succeededCount = jobsToRun.Count(job => job.Status == SuccessStatus);
        int skippedCount = jobsToRun.Count(job => job.Status == SkippedStatus);
        int failedCount = jobsToRun.Count(job => job.Status == FailedStatus);
        int completedCount = succeededCount + skippedCount + failedCount;

        CurrentFileProgressBar.IsIndeterminate = false;
        CurrentFileProgressBar.Value = completedCount == 0 ? 0 : 100;
        BatchProgressBar.Value = jobsToRun.Count == 0 ? 0 : 100;
        BatchProgressText.Text = $"{completedCount} / {jobsToRun.Count} 已完成";
        CurrentPptxText.Text = "当前 PPTX：批次已结束";
        CurrentDetailText.Text = "当前视频(已结束)";

        string infoBarMessage;
        InfoBarSeverity severity;

        if (failedCount > 0)
        {
            CurrentStatusText.Text = "当前状态：处理结束（有失败）";
            infoBarMessage = $"批次处理结束：成功 {succeededCount} 个，跳过 {skippedCount} 个，失败 {failedCount} 个。";
            severity = InfoBarSeverity.Error;
        }
        else if (succeededCount == 0 && skippedCount > 0)
        {
            CurrentStatusText.Text = "当前状态：处理结束";
            infoBarMessage = "本次没有生成新的输出文件。";
            severity = InfoBarSeverity.Warning;
        }
        else
        {
            CurrentStatusText.Text = "当前状态：空闲";
            infoBarMessage = $"全部待处理任务已完成：成功 {succeededCount} 个，未生成输出 {skippedCount} 个。";
            severity = InfoBarSeverity.Success;
        }

        ShowInfoBar(infoBarMessage, severity);
        App.WriteDiagnosticLog("开始处理", $"批次处理已结束。成功 {succeededCount} 个，跳过 {skippedCount} 个，失败 {failedCount} 个。");
    }

    private TimeSpan GetCurrentJobElapsed()
    {
        return _currentJobStopwatch?.Elapsed ?? TimeSpan.Zero;
    }

    private string BuildNoOutputReason(int totalMedia, int skippedCount, int failedCount, int alreadySatisfiedCount, int noVideoCount)
    {
        if (!_configHasVideoChanges)
        {
            return "当前未配置任何视频处理参数，因此未生成新的输出文件。";
        }

        if (totalMedia == 0 || (noVideoCount > 0 && noVideoCount == totalMedia))
        {
            return "PPTX 中未发现可处理视频，因此未生成新的输出文件。";
        }

        if (alreadySatisfiedCount > 0 && skippedCount == alreadySatisfiedCount && failedCount == 0)
        {
            return "所有视频都已满足目标条件，无需重新处理。";
        }

        if (failedCount > 0 && skippedCount == 0)
        {
            return "所有需要处理的视频都处理失败，因此未生成新的输出文件。";
        }

        if (failedCount > 0 && skippedCount > 0)
        {
            return $"未生成新的输出文件：失败 {failedCount} 个，跳过 {skippedCount} 个。";
        }

        return "未生成新的输出文件。";
    }

    private static string BuildQueueProgressText(int currentIndex, int totalCount)
    {
        if (totalCount <= 0)
        {
            return "准备中";
        }

        int displayIndex = Math.Clamp(currentIndex, 1, totalCount);
        return $"{displayIndex}/{totalCount} 视频";
    }

    private static string BuildCurrentVideoText(int currentIndex, int totalCount, double? filePercent)
    {
        if (totalCount <= 0)
        {
            return "当前视频(准备中)";
        }

        int displayIndex = Math.Clamp(currentIndex, 1, totalCount);
        string label = $"当前视频({displayIndex}/{totalCount})";

        if (!filePercent.HasValue)
        {
            return label;
        }

        return $"{label} {filePercent.Value:0.#}%";
    }

    private static string BuildSuccessDetailSummary(
        string inputPath,
        string outputPath,
        TimeSpan duration,
        int processedCount,
        int skippedCount,
        int failedCount,
        int alreadySatisfiedCount)
    {
        var lines = new List<string>
        {
            $"源文件：{inputPath}",
        };

        if (!string.IsNullOrWhiteSpace(outputPath))
        {
            lines.Add($"输出文件：{outputPath}");
        }

        lines.Add($"处理时间：{FormatDuration(duration)}");
        lines.Add($"结果：成功处理 {processedCount} 个视频，跳过 {skippedCount} 个，失败 {failedCount} 个。");

        if (alreadySatisfiedCount > 0)
        {
            lines.Add($"说明：其中 {alreadySatisfiedCount} 个视频已满足目标条件，未重复处理。");
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static string BuildSkippedDetailSummary(string inputPath, string outputPath, TimeSpan duration, string reason)
    {
        var lines = new List<string>
        {
            $"源文件：{inputPath}",
        };

        if (!string.IsNullOrWhiteSpace(outputPath))
        {
            lines.Add($"输出文件：{outputPath}");
        }

        lines.Add($"处理时间：{FormatDuration(duration)}");
        lines.Add($"原因：{reason}");
        return string.Join(Environment.NewLine, lines);
    }

    private static string BuildFailureDetailSummary(string inputPath, TimeSpan duration, string reason)
    {
        return string.Join(Environment.NewLine,
            $"源文件：{inputPath}",
            $"处理时间：{FormatDuration(duration)}",
            $"原因：{reason}");
    }

    private static string FormatDuration(TimeSpan duration)
    {
        if (duration < TimeSpan.FromSeconds(1))
        {
            return $"{Math.Max(duration.TotalMilliseconds, 1):0} 毫秒";
        }

        if (duration < TimeSpan.FromMinutes(1))
        {
            return $"{duration.TotalSeconds:0.#} 秒";
        }

        if (duration < TimeSpan.FromHours(1))
        {
            return $"{duration.Minutes} 分 {duration.Seconds} 秒";
        }

        return $"{(int)duration.TotalHours} 小时 {duration.Minutes} 分 {duration.Seconds} 秒";
    }

    private static IReadOnlyList<string> ParseSelectedFiles(string rawSelection)
    {
        string trimmed = rawSelection.TrimEnd('\0');

        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return Array.Empty<string>();
        }

        string[] parts = trimmed.Split('\0', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        if (parts.Length == 1)
        {
            return new[] { parts[0] };
        }

        string directory = parts[0];
        return parts.Skip(1)
            .Select(fileName => Path.Combine(directory, fileName))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string DescribeOpenFileDialogError(int errorCode)
    {
        return errorCode switch
        {
            0x3001 => "文件对话框创建失败。",
            0x3002 => "文件对话框初始化失败。",
            0x3003 => "文件对话框缓冲区太小。",
            0x3005 => "文件名无效。",
            0x3006 => "找不到指定实例。",
            0x3007 => "文件对话框模板无效。",
            0x3008 => "找不到资源。",
            0x3009 => "资源数据无效。",
            0x300A => "结构大小不正确。",
            0x300B => "未指定有效的钩子过程。",
            0x300C => "注册消息失败。",
            0x300D => "文件对话框模板未加载。",
            0x300E => "找不到指定菜单。",
            0x300F => "找不到指定类。",
            _ => string.Empty,
        };
    }

    private static string ExtractWorkerError(string stderr)
    {
        if (string.IsNullOrWhiteSpace(stderr))
        {
            return "native worker 已退出，但没有返回可读的错误信息。";
        }

        return stderr.Trim();
    }

    private static string ReadString(JsonElement root, string propertyName, string fallback)
    {
        if (!root.TryGetProperty(propertyName, out JsonElement value) || value.ValueKind != JsonValueKind.String)
        {
            return fallback;
        }

        return value.GetString() ?? fallback;
    }

    private static int ReadInt(JsonElement root, string propertyName, int fallback)
    {
        if (!root.TryGetProperty(propertyName, out JsonElement value) || !value.TryGetInt32(out int parsed))
        {
            return fallback;
        }

        return parsed;
    }

    private static double? ReadDouble(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out JsonElement value) || !value.TryGetDouble(out double parsed))
        {
            return null;
        }

        return parsed;
    }

    private static bool ReadBool(JsonElement root, string propertyName, bool fallback)
    {
        if (!root.TryGetProperty(propertyName, out JsonElement value) || value.ValueKind is not JsonValueKind.True and not JsonValueKind.False)
        {
            return fallback;
        }

        return value.GetBoolean();
    }

    [DllImport("comdlg32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool GetOpenFileName(ref OPENFILENAME openFileName);

    [DllImport("comdlg32.dll")]
    private static extern int CommDlgExtendedError();

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    private static extern nint SetWindowLongPtr(nint hWnd, int nIndex, nint dwNewLong);

    [DllImport("user32.dll")]
    private static extern nint CallWindowProc(nint lpPrevWndFunc, nint hWnd, uint msg, nuint wParam, nint lParam);

    [DllImport("user32.dll")]
    private static extern nint DefWindowProc(nint hWnd, uint msg, nuint wParam, nint lParam);

    [UnmanagedFunctionPointer(CallingConvention.Winapi)]
    private delegate nint WindowProcedure(nint hwnd, uint message, nuint wParam, nint lParam);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct OPENFILENAME
    {
        public int lStructSize;
        public IntPtr hwndOwner;
        public IntPtr hInstance;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpstrFilter;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpstrCustomFilter;
        public int nMaxCustFilter;
        public int nFilterIndex;
        public IntPtr lpstrFile;
        public int nMaxFile;
        public IntPtr lpstrFileTitle;
        public int nMaxFileTitle;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpstrInitialDir;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpstrTitle;
        public int Flags;
        public short nFileOffset;
        public short nFileExtension;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpstrDefExt;
        public IntPtr lCustData;
        public IntPtr lpfnHook;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? lpTemplateName;
        public IntPtr pvReserved;
        public int dwReserved;
        public int FlagsEx;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT
    {
        public int x;
        public int y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MINMAXINFO
    {
        public POINT ptReserved;
        public POINT ptMaxSize;
        public POINT ptMaxPosition;
        public POINT ptMinTrackSize;
        public POINT ptMaxTrackSize;
    }
}

































