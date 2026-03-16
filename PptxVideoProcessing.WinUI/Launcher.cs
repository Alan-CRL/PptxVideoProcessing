using System;
using System.Diagnostics;
using System.IO;

internal static class LauncherProgram
{
    [STAThread]
    private static int Main()
    {
        string appDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "App");
        string targetPath = Path.Combine(appDirectory, "PptxVideoProcessing.exe");

        if (!File.Exists(targetPath))
        {
            return 1;
        }

        Process process = Process.Start(new ProcessStartInfo
        {
            FileName = targetPath,
            WorkingDirectory = appDirectory,
            UseShellExecute = false
        });

        return process == null ? 1 : 0;
    }
}
