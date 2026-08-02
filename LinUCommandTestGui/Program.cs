using Avalonia;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace LinUCommandTestGui;

sealed class Program
{
    internal const string EditorProxyEnvironmentVariable = "LINU_QACLI_EDITOR_PROXY";

    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        if (TryRunQacliEditorProxy(args, out var editorExitCode))
        {
            Environment.ExitCode = editorExitCode;
            return;
        }

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            var ex = e.ExceptionObject as Exception;
            CrashLogger.Log("AppDomain.UnhandledException", ex);
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            CrashLogger.Log("TaskScheduler.UnobservedTaskException", e.Exception);
            e.SetObserved();
        };

        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex)
        {
            CrashLogger.Log("Fatal exception in Program.Main", ex);
            throw;
        }
    }

    private static bool TryRunQacliEditorProxy(string[] args, out int exitCode)
    {
        exitCode = 0;
        if (!OperatingSystem.IsLinux()
            || args.Length == 0
            || !string.Equals(
                Environment.GetEnvironmentVariable(EditorProxyEnvironmentVariable),
                "1",
                StringComparison.Ordinal))
        {
            return false;
        }

        const string terminalPath = "/usr/bin/gnome-terminal";
        const string editorPath = "/usr/bin/vi";
        if (!File.Exists(terminalPath) || !File.Exists(editorPath))
        {
            Console.Error.WriteLine("gnome-terminal または vi が見つかりません。");
            exitCode = 127;
            return true;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = terminalPath,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.ArgumentList.Add("--tab");
        startInfo.ArgumentList.Add("--");
        startInfo.ArgumentList.Add(editorPath);
        foreach (var argument in args)
        {
            startInfo.ArgumentList.Add(argument);
        }

        try
        {
            using var terminal = Process.Start(startInfo);
            if (terminal is null)
            {
                Console.Error.WriteLine("gnome-terminalを起動できませんでした。");
                exitCode = 127;
                return true;
            }

            terminal.WaitForExit();
            exitCode = terminal.ExitCode;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"gnome-terminalの起動に失敗しました: {ex.Message}");
            exitCode = 127;
        }

        return true;
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
#if DEBUG
            .WithDeveloperTools()
#endif
            .WithInterFont()
            .LogToTrace();
}
