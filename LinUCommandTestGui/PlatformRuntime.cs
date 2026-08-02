using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace LinUCommandTestGui;

/// <summary>
/// Centralizes behavior that must remain consistent across Windows and Linux.
/// This class deliberately has no Avalonia dependency so it can be tested on both CI runners.
/// </summary>
public static class PlatformRuntime
{
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
