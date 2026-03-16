using Microsoft.UI.Xaml;
using System.Text;

namespace PptxVideoProcessing.WinUI;

public partial class App : Application
{
    internal static readonly string ExecutableDirectory = Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;
    internal static readonly string RuntimeAssetsDirectory = ResolveRuntimeAssetsDirectory();
    private static readonly string DiagnosticLogPath = Path.Combine(ExecutableDirectory, "PptxVideoProcessing.WinUI.log");
    private Window? _window;

    public App()
    {
        UnhandledException += OnUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnCurrentDomainUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        InitializeComponent();
        WriteDiagnosticLog("应用启动", $"进程已启动。PID={Environment.ProcessId}");
    }

    private static string ResolveRuntimeAssetsDirectory()
    {
        string appDirectory = Path.Combine(ExecutableDirectory, "App");

        if (File.Exists(Path.Combine(appDirectory, "PptxVideoProcessing.Worker.exe")) ||
            File.Exists(Path.Combine(appDirectory, "ffmpeg.exe")) ||
            File.Exists(Path.Combine(appDirectory, "config.json")))
        {
            return appDirectory;
        }

        return ExecutableDirectory;
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        WriteDiagnosticLog("应用启动", "正在创建主窗口。");
        _window = new MainWindow();
        _window.Activate();
        WriteDiagnosticLog("应用启动", "主窗口已激活。");
    }

    private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
    {
        WriteCrashLog("WinUI 未处理异常", e.Exception);
    }

    private void OnCurrentDomainUnhandledException(object? sender, System.UnhandledExceptionEventArgs e)
    {
        WriteCrashLog("应用域未处理异常", e.ExceptionObject as Exception);
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        WriteCrashLog("后台任务未观察异常", e.Exception);
        e.SetObserved();
    }

    internal static void WriteDiagnosticLog(string source, string message)
    {
        AppendLog(source, message);
    }

    internal static void WriteCrashLog(string source, Exception? exception)
    {
        string details = exception is null
            ? "<无异常对象>"
            : exception.ToString();

        AppendLog(source, details);
    }

    private static void AppendLog(string source, string details)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(DiagnosticLogPath) ?? ExecutableDirectory);

            var builder = new StringBuilder();
            builder.AppendLine("====================");
            builder.AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            builder.AppendLine(source);
            builder.AppendLine(details);

            File.AppendAllText(DiagnosticLogPath, builder.ToString(), Encoding.UTF8);
        }
        catch
        {
            // Swallow logging failures to avoid cascading crashes.
        }
    }
}

