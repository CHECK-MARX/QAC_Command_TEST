using System;
using System.IO;
using System.Text;

namespace LinUCommandTestGui;

internal static class CrashLogger
{
    private static readonly object Sync = new();

    public static string GetLogPath()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (string.IsNullOrWhiteSpace(home))
        {
            home = "/tmp";
        }

        return Path.Combine(home, ".linu_gui_crash.log");
    }

    public static void Log(string context, Exception? ex = null)
    {
        try
        {
            var path = GetLogPath();
            var sb = new StringBuilder();
            sb.AppendLine("======================================");
            sb.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            sb.AppendLine($"Context  : {context}");
            if (ex is not null)
            {
                sb.AppendLine($"Type     : {ex.GetType().FullName}");
                sb.AppendLine($"Message  : {ex.Message}");
                sb.AppendLine("StackTrace:");
                sb.AppendLine(ex.ToString());
            }

            lock (Sync)
            {
                File.AppendAllText(path, sb.ToString(), Encoding.UTF8);
            }
        }
        catch
        {
            // Never throw from logger.
        }
    }
}
