using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace LinUCommandTestGui;

/// <summary>
/// Centralizes behavior that must remain consistent across Windows and Linux.
/// This class deliberately has no Avalonia dependency so it can be tested on both CI runners.
/// </summary>
public static class PlatformRuntime
{
    private static readonly Regex RetryableQacliLicenseMessageRegex = new(
        "ライセンスが(拒否|欠如)|ライセンス.*(不足|利用できません)|license.*(denied|refused|missing|unavailable)|communications error with license server",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    // Bash runs this file through BASH_ENV for the test script. Keep qacli output live:
    // qacli can ask for input before it exits, and the GUI must see that prompt.
    public const string LinuxQacliLicenseRetryScript = """
        # GUI runtime guard for transient Perforce QAC license contention.
        # This file is generated automatically; edit PlatformRuntime.cs instead.
        qacli() {
          local qacli_attempt=1
          local qacli_max_attempts=4
          local qacli_retry_wait_seconds=30
          local qacli_rc=0
          local qacli_output_file

          while [ "$qacli_attempt" -le "$qacli_max_attempts" ]; do
            qacli_output_file="$(mktemp "${TMPDIR:-/tmp}/linu-qacli.XXXXXX")" || return 1
            "$LINU_REAL_QACLI" "$@" 2>&1 | tee "$qacli_output_file"
            qacli_rc=${PIPESTATUS[0]}

            if grep -Eiq 'ライセンスが(拒否|欠如)|ライセンス.*(不足|利用できません)|license.*(denied|refused|missing|unavailable)|communications error with license server' "$qacli_output_file"; then
              rm -f "$qacli_output_file"
              if [ "$qacli_attempt" -lt "$qacli_max_attempts" ]; then
                printf '[GUI-LICENSE-RETRY] ライセンス確保待ち: %s秒後に再試行します (%s/%s)\n' \
                  "$qacli_retry_wait_seconds" "$qacli_attempt" "$qacli_max_attempts"
                sleep "$qacli_retry_wait_seconds"
                qacli_attempt=$((qacli_attempt + 1))
                continue
              fi
              printf '[ERR] qacli license retry exhausted after %s attempts.\n' "$qacli_max_attempts"
              if [ "$qacli_rc" -eq 0 ]; then qacli_rc=1; fi
              return "$qacli_rc"
            fi

            rm -f "$qacli_output_file"
            return "$qacli_rc"
          done
        }
        export -f qacli
        """;

    public static bool IsRetryableQacliLicenseMessage(string line)
    {
        return RetryableQacliLicenseMessageRegex.IsMatch(line);
    }

    public const string ValidateAuthTimeoutEnvironmentVariable = "LINU_VAL_AUTH_TIMEOUT_SECONDS";
    public const int DefaultValidateAuthTimeoutSeconds = 60;
    public const int MinimumValidateAuthTimeoutSeconds = 10;
    public const int MaximumValidateAuthTimeoutSeconds = 600;

    private static readonly Lazy<string> BuildIdentityValue = new(CreateBuildIdentity);

    public static string BuildIdentity => BuildIdentityValue.Value;

    public static TimeSpan GetValidateAuthTimeout(Func<string, string?>? readEnvironmentVariable = null)
    {
        readEnvironmentVariable ??= Environment.GetEnvironmentVariable;
        var configuredValue = readEnvironmentVariable(ValidateAuthTimeoutEnvironmentVariable);
        if (!int.TryParse(configuredValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out var seconds))
        {
            seconds = DefaultValidateAuthTimeoutSeconds;
        }

        return TimeSpan.FromSeconds(Math.Clamp(
            seconds,
            MinimumValidateAuthTimeoutSeconds,
            MaximumValidateAuthTimeoutSeconds));
    }

    public static Encoding GetQacliOutputEncoding()
    {
        return GetQacliOutputEncoding(OperatingSystem.IsWindows());
    }

    public static Encoding GetQacliOutputEncoding(bool isWindows)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return isWindows
            ? Encoding.GetEncoding(932)
            : new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
    }

    public static ProcessStartInfo CreateRedirectedProcessStartInfo(
        string fileName,
        IReadOnlyList<string> arguments,
        string workingDirectory,
        Encoding outputEncoding)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentException.ThrowIfNullOrWhiteSpace(workingDirectory);
        ArgumentNullException.ThrowIfNull(outputEncoding);

        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            WorkingDirectory = workingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = outputEncoding,
            StandardErrorEncoding = outputEncoding
        };

        foreach (var argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        return startInfo;
    }

    public static async Task<RedirectedProcessResult> RunRedirectedProcessAsync(
        string fileName,
        IReadOnlyList<string> arguments,
        string workingDirectory,
        Encoding outputEncoding,
        TimeSpan timeout)
    {
        using var process = new Process
        {
            StartInfo = CreateRedirectedProcessStartInfo(
                fileName,
                arguments,
                workingDirectory,
                outputEncoding)
        };

        if (!process.Start())
        {
            throw new InvalidOperationException($"Failed to start precheck command: {fileName}");
        }

        var stdOutTask = process.StandardOutput.ReadToEndAsync();
        var stdErrTask = process.StandardError.ReadToEndAsync();
        var waitTask = process.WaitForExitAsync();
        var timeoutTask = Task.Delay(timeout);

        var completedTask = await Task.WhenAny(waitTask, timeoutTask);
        var timedOut = !ReferenceEquals(completedTask, waitTask);
        if (timedOut)
        {
            try
            {
                process.Kill(entireProcessTree: true);
            }
            catch (InvalidOperationException)
            {
                // The process exited between the timeout and the kill request.
            }

            await waitTask;
        }

        return new RedirectedProcessResult(
            process.ExitCode,
            await stdOutTask,
            await stdErrTask,
            timedOut);
    }

    private static string CreateBuildIdentity()
    {
        var assembly = typeof(PlatformRuntime).Assembly;
        var informationalVersion = assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?
            .InformationalVersion;
        var version = informationalVersion?.Split('+', 2)[0]
            ?? assembly.GetName().Version?.ToString(3)
            ?? "unknown";
        var metadata = informationalVersion?.Split('+', 2).ElementAtOrDefault(1) ?? string.Empty;
        var fallbackCommit = new string(metadata
            .TakeWhile(Uri.IsHexDigit)
            .Take(8)
            .ToArray());
        var declaredRevision = assembly
            .GetCustomAttributes<AssemblyMetadataAttribute>()
            .FirstOrDefault(attribute => string.Equals(
                attribute.Key,
                "BuildSourceRevision",
                StringComparison.Ordinal))?
            .Value;
        var revision = string.Equals(declaredRevision, "local", StringComparison.OrdinalIgnoreCase)
            ? "local"
            : new string((declaredRevision ?? fallbackCommit)
                .TakeWhile(Uri.IsHexDigit)
                .Take(8)
                .ToArray());
        var runtimeIdentifier = RuntimeInformation.RuntimeIdentifier;

        return string.IsNullOrEmpty(revision)
            ? $"v{version} · {runtimeIdentifier}"
            : $"v{version} · {revision} · {runtimeIdentifier}";
    }
}

public sealed record RedirectedProcessResult(
    int ExitCode,
    string StdOut,
    string StdErr,
    bool TimedOut);
