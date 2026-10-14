using System.Diagnostics;
using Xunit;
using LinUCommandTestGui.ViewModels;

namespace LinUCommandTestGui.Tests;

public sealed class LinuxTestRuntimeTests
{
    [Fact]
    public void FifteenTerminalDiagnosticsDoNotBecomeQacFailures()
    {
        string[] terminalLines =
        [
            "libEGL warning: failed to get driver name for fd -1",
            "libEGL warning: MESA-LOADER: failed to retrieve device information",
            "libEGL warning: failed to get driver name for fd -1",
            "MESA: error: ZINK: failed to choose pdev",
            "libEGL warning: egl: failed to create dri2 screen"
        ];
        var threeLoops = Enumerable.Range(0, 3).SelectMany(_ => terminalLines).ToArray();
        Assert.Equal(15, threeLoops.Count(MainWindowViewModel.IsFailureOutputLine));
        Assert.Equal(0, threeLoops.Count(line => MainWindowViewModel.IsFailureOutputLine("[TERMINAL-DIAGNOSTIC] " + line)));
    }

    [Theory]
    [InlineData("エラーメッセージ: 前のスナップショットを処理中です。")]
    [InlineData("[ERR] 編集用端末の実行に失敗しました (exit=127)")]
    [InlineData("[ERR] 編集コマンドが失敗しました (exit=7)")]
    [InlineData("1 成功および、1 失敗。")]
    [InlineData("fatal: license unavailable")]
    public void RealFailuresStillFail(string line)
    {
        Assert.True(MainWindowViewModel.IsFailureOutputLine(line));
    }

    [Fact]
    public void EnterIsSentOnlyOnceAndNeverForAnOldPrompt()
    {
        var gate = new PromptConfirmationGate();
        Assert.False(gate.TryConsume(gate.Generation));
        gate.Observe();
        var first = gate.Generation;
        Assert.True(gate.TryConsume(first));
        Assert.False(gate.TryConsume(first));
        gate.Observe();
        Assert.False(gate.TryConsume(first));
        Assert.True(gate.TryConsume(gate.Generation));
        gate.Observe();
        var last = gate.Generation;
        gate.Clear();
        Assert.False(gate.TryConsume(last));
    }

    [Fact]
    public void OriginalScriptWinsAndGeneratedScriptIsNeverUsed()
    {
        var root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
        Directory.CreateDirectory(root);
        try
        {
            File.WriteAllText(Path.Combine(root, "testloop.generated.sh"), "exit 0");
            Assert.Null(LinuxTestRuntime.FindOriginalScript(root));
            var original = Path.Combine(root, "testloop.sh");
            File.WriteAllText(original, "exit 1");
            Assert.Equal(original, LinuxTestRuntime.FindOriginalScript(root));
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task EditorCanReadTerminalEvenWhenParentPipesItsInput()
    {
        if (!OperatingSystem.IsLinux() || !File.Exists("/usr/bin/vi")) return;
        var root = Path.Combine(Path.GetTempPath(), "qac-editor-" + Guid.NewGuid());
        Directory.CreateDirectory(root);
        try
        {
            var helper = Path.Combine(root, "editor.sh");
            var file = Path.Combine(root, "test.acf");
            await File.WriteAllTextAsync(helper, LinuxTestRuntime.EditorScript);
            await File.WriteAllTextAsync(file, "unchanged\n");
            File.SetUnixFileMode(helper, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            var start = new ProcessStartInfo("/usr/bin/script")
            {
                RedirectStandardInput = true, RedirectStandardOutput = true,
                RedirectStandardError = true, UseShellExecute = false
            };
            start.Environment["TERM"] = "xterm";
            foreach (var arg in new[] { "-q", "-e", "-E", "never", "-c", $"sleep 2 | '{helper}' '{file}'", "/dev/null" })
                start.ArgumentList.Add(arg);
            using var process = Process.Start(start)!;
            try
            {
                var buffer = new char[1];
                Assert.Equal(1, await process.StandardOutput.ReadAsync(buffer.AsMemory()).AsTask().WaitAsync(TimeSpan.FromSeconds(5)));
                await process.StandardInput.WriteAsync("\u001b:q\r");
                await process.StandardInput.FlushAsync();
                var output = process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(8));
                await output;
                Assert.Equal(0, process.ExitCode);
                Assert.Equal("unchanged\n", await File.ReadAllTextAsync(file));
            }
            finally { if (!process.HasExited) process.Kill(entireProcessTree: true); }
        }
        finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task BridgePreservesCommandsSubstitutionExitStatusAndLiveOutput()
    {
        if (!OperatingSystem.IsLinux()) return;
        var root = Path.Combine(Path.GetTempPath(), "qac test 日本語 " + Guid.NewGuid());
        Directory.CreateDirectory(root);
        try
        {
            var bridge = Path.Combine(root, "bridge.sh");
            var fake = Path.Combine(root, "fake-qacli");
            var trace = Path.Combine(root, "trace");
            await File.WriteAllTextAsync(bridge, LinuxTestRuntime.ShellBridge);
            await File.WriteAllTextAsync(fake, """
                #!/bin/bash
                printf '%s\n' "$*" >> "$TRACE"
                case "$1" in
                  early) printf 'EARLY\n'; sleep 2; printf 'LATE\n' ;;
                  fail) exit 7 ;;
                  project) exit "${EDIT_RESULT:-7}" ;;
                  value) printf '%s' "$2" ;;
                  tty) test -t 0 && test -t 1 && test -t 2 || exit 42 ;;
                esac
                """);
            File.SetUnixFileMode(fake, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            var start = new ProcessStartInfo("bash")
            {
                RedirectStandardOutput = true, RedirectStandardError = true,
                RedirectStandardInput = true, UseShellExecute = false
            };
            start.Environment["BASH_ENV"] = bridge;
            start.Environment["LINU_REAL_QACLI"] = fake;
            start.Environment["TRACE"] = trace;
            start.ArgumentList.Add("-c");
            start.ArgumentList.Add("qacli early; value=$(qacli value '日本語 空白'); [ \"$value\" = '日本語 空白' ] || exit 99; qacli fail; rc=$?; printf 'STATUS=%s\\n' \"$rc\"; exit \"$rc\"");
            using (var process = Process.Start(start)!)
            {
                Assert.Equal("EARLY", await process.StandardOutput.ReadLineAsync().WaitAsync(TimeSpan.FromSeconds(5)));
                Assert.False(process.HasExited); // Output arrives while the command is still running.
                var remainder = await process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
                Assert.Equal(7, process.ExitCode);
                Assert.Contains("LATE", remainder);
                Assert.Contains("STATUS=7", remainder);
            }
            Assert.Equal(new[] { "early", "value 日本語 空白", "fail" }, await File.ReadAllLinesAsync(trace));

            // The GUI uses script(1): ordinary commands must retain all three terminal fds.
            start.FileName = "/usr/bin/script";
            start.ArgumentList.Clear();
            foreach (var arg in new[] { "-q", "-e", "-E", "never", "-c", "bash -c 'qacli tty'", "/dev/null" })
                start.ArgumentList.Add(arg);
            using var terminal = Process.Start(start)!;
            terminal.StandardInput.Close();
            await terminal.StandardOutput.ReadToEndAsync();
            await terminal.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(0, terminal.ExitCode);

            // Simulate a terminal that reports success even if its child fails.
            var terminalStub = Path.Combine(root, "ptyxis");
            await File.WriteAllTextAsync(terminalStub, "#!/bin/sh\nprintf 'libEGL warning: failed to get driver name for fd -1\\n' >&2\nwhile [ \"$1\" != -- ]; do shift; done\nshift\n\"$@\"\nexit 0\n");
            File.SetUnixFileMode(terminalStub, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            start.FileName = "bash";
            start.Environment["PATH"] = root + ":" + Environment.GetEnvironmentVariable("PATH");
            start.ArgumentList.Clear();
            start.ArgumentList.Add("-c");
            start.ArgumentList.Add("qacli project config-file -P '日本語 空白' --edit acf");
            using var editor = Process.Start(start)!;
            var editorOutput = await editor.StandardOutput.ReadToEndAsync();
            await editor.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(7, editor.ExitCode);
            Assert.Contains("[EDITOR]", editorOutput);
            Assert.Equal("project config-file -P 日本語 空白 --edit acf", (await File.ReadAllLinesAsync(trace))[^1]);
            Assert.Contains("[TERMINAL-DIAGNOSTIC] libEGL warning:", editorOutput);
            Assert.Contains("[ERR]", await editor.StandardError.ReadToEndAsync());

            start.Environment["EDIT_RESULT"] = "0";
            using var successfulEditor = Process.Start(start)!;
            var successOutput = await successfulEditor.StandardOutput.ReadToEndAsync();
            await successfulEditor.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(0, successfulEditor.ExitCode);
            Assert.Contains("[TERMINAL-DIAGNOSTIC]", successOutput);
            Assert.DoesNotContain(successOutput.Split('\n'), MainWindowViewModel.IsFailureOutputLine);

            // A terminal that cannot launch must still fail, even when all diagnostics are tagged.
            await File.WriteAllTextAsync(terminalStub, "#!/bin/sh\necho 'failed to open display' >&2\nexit 127\n");
            using var failedTerminal = Process.Start(start)!;
            await failedTerminal.StandardOutput.ReadToEndAsync();
            var launchError = await failedTerminal.StandardError.ReadToEndAsync();
            await failedTerminal.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(127, failedTerminal.ExitCode);
            Assert.Contains(launchError.Split('\n'), MainWindowViewModel.IsFailureOutputLine);


        }
        finally { Directory.Delete(root, true); }
    }
}
