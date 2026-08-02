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

    private static (string FileName, string[] Arguments) CreateShellCommand(string command)
    {
        return OperatingSystem.IsWindows()
            ? ("cmd.exe", ["/d", "/c", command])
            : ("/bin/sh", ["-c", command]);
    }
}
