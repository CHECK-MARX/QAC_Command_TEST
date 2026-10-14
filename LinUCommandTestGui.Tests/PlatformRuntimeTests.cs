using System.Diagnostics;
using System.Text;
using Xunit;

namespace LinUCommandTestGui.Tests;

public sealed class PlatformRuntimeTests
{
    [Fact]
    public void ValidateAuthTimeoutDefaultsToSixtySeconds()
    {
        var timeout = PlatformRuntime.GetValidateAuthTimeout(_ => null);

        Assert.Equal(TimeSpan.FromSeconds(60), timeout);
    }

    [Theory]
    [InlineData("9", 10)]
    [InlineData("45", 45)]
    [InlineData("900", 600)]
    [InlineData("invalid", 60)]
    public void ValidateAuthTimeoutIsValidatedAndClamped(string value, int expectedSeconds)
    {
        var timeout = PlatformRuntime.GetValidateAuthTimeout(_ => value);

        Assert.Equal(TimeSpan.FromSeconds(expectedSeconds), timeout);
    }

    [Fact]
    public void QacliEncodingUsesCp932OnWindows()
    {
        var encoding = PlatformRuntime.GetQacliOutputEncoding(isWindows: true);

        Assert.Equal(932, encoding.CodePage);
    }

    [Fact]
    public void QacliEncodingUsesUtf8OnLinux()
    {
        var encoding = PlatformRuntime.GetQacliOutputEncoding(isWindows: false);

        Assert.Equal(Encoding.UTF8.CodePage, encoding.CodePage);
        Assert.Empty(encoding.GetPreamble());
    }

    [Fact]
    public void PrecheckArgumentsAreKeptAsSeparateValues()
    {
        string[] arguments =
        [
            "auth",
            "--username",
            "user name",
            "--password",
            "quote\" and & shell characters",
            "--url",
            "http://127.0.0.1:8282"
        ];

        var startInfo = PlatformRuntime.CreateRedirectedProcessStartInfo(
            "qacli",
            arguments,
            Environment.CurrentDirectory,
            PlatformRuntime.GetQacliOutputEncoding());

        Assert.Equal(arguments, startInfo.ArgumentList);
        Assert.Equal(string.Empty, startInfo.Arguments);
        Assert.False(startInfo.UseShellExecute);
        Assert.True(startInfo.RedirectStandardOutput);
        Assert.True(startInfo.RedirectStandardError);
    }

    [Fact]
    public void BuildIdentityIdentifiesVersionAndRuntime()
    {
        Assert.StartsWith("v", PlatformRuntime.BuildIdentity);
        Assert.Matches(@" · (local|[0-9a-fA-F]{8}) · ", PlatformRuntime.BuildIdentity);
        Assert.Contains(System.Runtime.InteropServices.RuntimeInformation.RuntimeIdentifier, PlatformRuntime.BuildIdentity);
    }

    [Fact]
    public async Task RedirectedProcessRunsAndCapturesOutputOnCurrentPlatform()
    {
        var (fileName, arguments) = CreateShellCommand("echo platform-check");

        var result = await PlatformRuntime.RunRedirectedProcessAsync(
            fileName,
            arguments,
            Environment.CurrentDirectory,
            PlatformRuntime.GetQacliOutputEncoding(),
            TimeSpan.FromSeconds(5));

        Assert.False(result.TimedOut);
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("platform-check", result.StdOut);
    }

    [Fact]
    public async Task RedirectedProcessTerminatesAfterTimeoutOnCurrentPlatform()
    {
        var command = OperatingSystem.IsWindows()
            ? "ping -n 4 127.0.0.1 >NUL"
            : "sleep 3";
        var (fileName, arguments) = CreateShellCommand(command);

        var result = await PlatformRuntime.RunRedirectedProcessAsync(
            fileName,
            arguments,
            Environment.CurrentDirectory,
            PlatformRuntime.GetQacliOutputEncoding(),
            TimeSpan.FromMilliseconds(100));

        Assert.True(result.TimedOut);
    }

    [Fact]
    public async Task LinuxQacliRetryWrapperForwardsInteractivePromptBeforeQacliExits()
    {
        if (!OperatingSystem.IsLinux())
        {
            return;
        }

        var testDirectory = Path.Combine(Path.GetTempPath(), $"linu-qacli-test-{Guid.NewGuid():N}");
        Directory.CreateDirectory(testDirectory);
        var environmentPath = Path.Combine(testDirectory, "qacli-license-retry.sh");
        var qacliPath = Path.Combine(testDirectory, "mock-qacli.sh");
        try
        {
            await File.WriteAllTextAsync(environmentPath, PlatformRuntime.LinuxQacliLicenseRetryScript);
            await File.WriteAllTextAsync(qacliPath, "#!/usr/bin/env bash\nprintf 'Press [Enter] to continue.'\nIFS= read -r answer\nprintf 'completed\\n'\n");
            File.SetUnixFileMode(qacliPath,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);

            using var process = new Process
            {
                StartInfo = new ProcessStartInfo("bash", "-c qacli")
                {
                    WorkingDirectory = testDirectory,
                    UseShellExecute = false,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                }
            };
            process.StartInfo.Environment["BASH_ENV"] = environmentPath;
            process.StartInfo.Environment["LINU_REAL_QACLI"] = qacliPath;
            Assert.True(process.Start());
            try
            {
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                var prompt = new char["Press [Enter] to continue.".Length];
                var received = 0;
                while (received < prompt.Length)
                {
                    var count = await process.StandardOutput.ReadAsync(prompt.AsMemory(received), timeout.Token);
                    Assert.True(count > 0, "qacli exited before displaying the prompt.");
                    received += count;
                }

                Assert.Equal("Press [Enter] to continue.", new string(prompt));
                Assert.False(process.HasExited);
                await process.StandardInput.WriteLineAsync();
                await process.StandardInput.FlushAsync();
                await process.WaitForExitAsync(timeout.Token);
                Assert.Equal(0, process.ExitCode);
                Assert.Contains("completed", await process.StandardOutput.ReadToEndAsync());
            }
            finally
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                    await process.WaitForExitAsync();
                }
            }
        }
        finally
        {
            Directory.Delete(testDirectory, recursive: true);
        }
    }

    [Theory]
    [InlineData("license denied", true)]
    [InlineData("communications error with license server", true)]
    [InlineData("ライセンスが拒否されました", true)]
    [InlineData("qacli command failed", false)]
    [InlineData("[ERR] qacli license retry exhausted after 4 attempts.", false)]
    public void RetryableLicenseMessagesAreRecognizedSeparatelyFromFinalFailures(string line, bool expected)
    {
        Assert.Equal(expected, PlatformRuntime.IsRetryableQacliLicenseMessage(line));
    }

    private static (string FileName, string[] Arguments) CreateShellCommand(string command)
    {
        return OperatingSystem.IsWindows()
            ? ("cmd.exe", ["/d", "/c", command])
            : ("/bin/sh", ["-c", command]);
    }
}
