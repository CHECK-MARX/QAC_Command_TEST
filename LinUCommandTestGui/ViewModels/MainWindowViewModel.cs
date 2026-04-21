using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace LinUCommandTestGui.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    public ObservableCollection<ProjectPairItemViewModel> CandidatePairs { get; } = new();
    public ObservableCollection<LiveLogEntry> LogLines { get; } = new();
    public ObservableCollection<LanguageOptionViewModel> LanguageOptions { get; } =
    [
        new LanguageOptionViewModel("ja", "日本語"),
        new LanguageOptionViewModel("en", "English")
    ];
    public LocalizationTexts T { get; } = new();

    [ObservableProperty]
    private string linURootPath = DetectDefaultLinURoot();

    [ObservableProperty]
    private string commandTestDirectoryPath = DetectDefaultCommandTestDirectory();

    [ObservableProperty]
    private string allITablePath = string.Empty;

    [ObservableProperty]
    private string allJTablePath = string.Empty;

    [ObservableProperty]
    private string activeITablePath = string.Empty;

    [ObservableProperty]
    private string activeJTablePath = string.Empty;

    [ObservableProperty]
    private string runtimeDirectoryPath = string.Empty;

    [ObservableProperty]
    private string outputLogDirectoryPath = string.Empty;

    [ObservableProperty]
    private string qafRoot = string.Empty;

    [ObservableProperty]
    private string qacliBinPath = string.Empty;

    [ObservableProperty]
    private string testRoot = string.Empty;

    [ObservableProperty]
    private string comQac = string.Empty;

    [ObservableProperty]
    private string comQacpp = string.Empty;

    [ObservableProperty]
    private string comRcma = string.Empty;

    [ObservableProperty]
    private string comMta = string.Empty;

    [ObservableProperty]
    private string comName = string.Empty;

    [ObservableProperty]
    private string comData = string.Empty;

    [ObservableProperty]
    private string qavServer = string.Empty;

    [ObservableProperty]
    private string qavUser = string.Empty;

    [ObservableProperty]
    private string qavPass = string.Empty;

    [ObservableProperty]
    private string valServer = string.Empty;

    [ObservableProperty]
    private string valUser = string.Empty;

    [ObservableProperty]
    private string valPass = string.Empty;

    [ObservableProperty]
    private bool autoConfirmPrompts = true;

    [ObservableProperty]
    private int promptDelayMilliseconds = 1300;

    [ObservableProperty]
    private bool stopOnFirstError;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(StartTestCommand))]
    [NotifyCanExecuteChangedFor(nameof(StopTestCommand))]
    private bool isRunning;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(StartTestCommand))]
    [NotifyCanExecuteChangedFor(nameof(StopTestCommand))]
    private bool isPaused;

    [ObservableProperty]
    private string statusText = "Idle";

    [ObservableProperty]
    private string runResult = "Not started";

    [ObservableProperty]
    private string lastError = string.Empty;

    [ObservableProperty]
    private string currentRunLogPath = string.Empty;

    [ObservableProperty]
    private int configuredLoopCount;

    [ObservableProperty]
    private int processedSummaryCount;

    [ObservableProperty]
    private int totalSuccessInSummaries;

    [ObservableProperty]
    private int totalFailureInSummaries;

    [ObservableProperty]
    private int autoEnterCount;

    [ObservableProperty]
    private int selectedPairCount;

    [ObservableProperty]
    private string liveOutputText = string.Empty;

    [ObservableProperty]
    private bool autoScrollLiveOutput = true;

    [ObservableProperty]
    private int runningProcessId;

    [ObservableProperty]
    private string elapsedTime = "00:00:00";

    [ObservableProperty]
    private double progressPercentage;

    [ObservableProperty]
    private string progressText = "0 / 0 (0.0%)";

    [ObservableProperty]
    private string estimatedFinishTime = "-";

    [ObservableProperty]
    private string estimatedRemaining = "-";

    [ObservableProperty]
    private int loopStartedCount;

    [ObservableProperty]
    private int loopCompletedCount;

    [ObservableProperty]
    private int loopFailedCount;

    [ObservableProperty]
    private int errorLineCount;

    [ObservableProperty]
    private string verdictText = "Not started";

    [ObservableProperty]
    private string verdictReason = "-";

    [ObservableProperty]
    private string outputActivity = "-";

    [ObservableProperty]
    private string lastOutputTime = "-";

    [ObservableProperty]
    private int lastOutputAgeSeconds;

    [ObservableProperty]
    private LanguageOptionViewModel? selectedLanguage;

    private DateTime _runStartedUtc;
    private Process? _runningProcess;
    private bool _stopRequestedByUser;
    private readonly SemaphoreSlim _logWriteSemaphore = new(1, 1);
    private static readonly string[] TrackedEnvironmentKeys =
    [
        "QAF_ROOT",
        "QACLI_BIN",
        "TEST_ROOT",
        "COM_QAC",
        "COM_QACPP",
        "COM_RCMA",
        "COM_MTA",
        "COM_NAME",
        "COM_DATA",
        "QAV_SERVER",
        "QAV_USER",
        "QAV_PASS",
        "VAL_SERVER",
        "VAL_USER",
        "VAL_PASS"
    ];

    public MainWindowViewModel()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        T.SetLanguage("ja");
        SelectedLanguage = LanguageOptions.FirstOrDefault();
        InitializePathsFromRoot();

        if (Directory.Exists(LinURootPath))
        {
            TestRoot = LinURootPath;
        }

        CandidatePairs.CollectionChanged += (_, _) => RefreshSelectedPairCount();
        _ = ReloadPairsAsync();
    }

    partial void OnSelectedLanguageChanged(LanguageOptionViewModel? value)
    {
        if (value is null)
        {
            return;
        }

        T.SetLanguage(value.Code);
    }

    partial void OnLinURootPathChanged(string value)
    {
        InitializePathsFromRoot();
    }

    public void ReportUnhandledException(string message)
    {
        SetError(message);
    }

    private void InitializePathsFromRoot()
    {
        var masterSettingsDir = Path.Combine(LinURootPath, "master_settings");
        var allDir = Path.Combine(masterSettingsDir, "all");

        var sjisI = Path.Combine(masterSettingsDir, "JISX0208_SJIS_I.txt");
        var sjisJ = Path.Combine(masterSettingsDir, "JISX0208_SJIS_J.txt");
        var utf8I = Path.Combine(masterSettingsDir, "JISX0208_UTF8_I.txt");
        var utf8J = Path.Combine(masterSettingsDir, "JISX0208_UTF8_J.txt");

        if (File.Exists(sjisI) || File.Exists(sjisJ))
        {
            AllITablePath = Path.Combine(allDir, "JISX0208_SJIS_I.txt");
            AllJTablePath = Path.Combine(allDir, "JISX0208_SJIS_J.txt");
            ActiveITablePath = sjisI;
            ActiveJTablePath = sjisJ;
        }
        else
        {
            AllITablePath = Path.Combine(allDir, "JISX0208_UTF8_I.txt");
            AllJTablePath = Path.Combine(allDir, "JISX0208_UTF8_J.txt");
            ActiveITablePath = utf8I;
            ActiveJTablePath = utf8J;
        }

        RuntimeDirectoryPath = Path.Combine(LinURootPath, ".gui_runtime");
        OutputLogDirectoryPath = Path.Combine(LinURootPath, "gui_logs");
    }

    [RelayCommand]
    private async Task LoadFromCommandTestDirectoryAsync()
    {
        try
        {
            var rootCandidate = CommandTestDirectoryPath.Trim();
            if (string.IsNullOrWhiteSpace(rootCandidate))
            {
                throw new InvalidOperationException("Command test directory is empty.");
            }

            if (Path.GetFileName(rootCandidate).Equals("Lin_U", StringComparison.OrdinalIgnoreCase))
            {
                LinURootPath = rootCandidate;
                CommandTestDirectoryPath = Directory.GetParent(rootCandidate)?.FullName ?? rootCandidate;
            }
            else
            {
                var nested = Path.Combine(rootCandidate, "Lin_U");
                LinURootPath = Directory.Exists(nested) ? nested : rootCandidate;
            }

            TestRoot = LinURootPath;
            await ReloadPairsAsync();
            var envPath = await LoadEnvironmentFromSourceScriptAsync();
            if (envPath is null)
            {
                StatusText = $"Loaded command test directory: {LinURootPath}";
            }
            else
            {
                StatusText = $"Loaded command test directory and environment: {Path.GetFileName(envPath)}";
            }
        }
        catch (Exception ex)
        {
            SetError($"Failed to load command test directory: {ex.Message}");
        }
    }

    [RelayCommand]
    private async Task SaveEnvironmentToSourceScriptAsync()
    {
        try
        {
            var path = ResolveEnvironmentScriptPathForSave();
            EnsureDirectoryForFile(path);

            if (path.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
            {
                await SaveEnvironmentToBatchScriptAsync(path);
            }
            else
            {
                await SaveEnvironmentToShellScriptAsync(path);
            }

            StatusText = $"Saved GUI environment values to {Path.GetFileName(path)}.";
            AppendLog($"[GUI] {Path.GetFileName(path)} updated from GUI values.", false, false);
        }
        catch (Exception ex)
        {
            SetError($"Failed to save environment to script: {ex.Message}");
        }
    }

    [RelayCommand]
    private async Task ReloadPairsAsync()
    {
        try
        {
            var iLines = ReadLinesPreserveContent(ActiveITablePath);
            var jLines = ReadLinesPreserveContent(ActiveJTablePath);
            var count = Math.Max(iLines.Count, jLines.Count);

            CandidatePairs.Clear();
            for (var i = 0; i < count; i++)
            {
                var iValue = i < iLines.Count ? iLines[i] : string.Empty;
                var jValue = i < jLines.Count ? jLines[i] : string.Empty;
                var isSelected = !string.IsNullOrWhiteSpace(iValue) && !string.IsNullOrWhiteSpace(jValue);
                CandidatePairs.Add(CreatePairItem(i + 1, iValue, jValue, isSelected));
            }

            if (CandidatePairs.Count == 0)
            {
                CandidatePairs.Add(CreatePairItem(1, string.Empty, string.Empty, true));
            }

            RefreshSelectedPairCount();
            StatusText = $"Loaded {CandidatePairs.Count} test pairs.";
        }
        catch (Exception ex)
        {
            SetError($"Failed to load pairs: {ex.Message}");
        }

        await Task.CompletedTask;
    }

    [RelayCommand]
    private async Task SelectFirstThreeAsync()
    {
        for (var i = 0; i < CandidatePairs.Count; i++)
        {
            CandidatePairs[i].IsSelected = i < 3;
        }

        RefreshSelectedPairCount();
        await Task.CompletedTask;
    }

    [RelayCommand]
    private async Task AddPairRowAsync()
    {
        CandidatePairs.Add(CreatePairItem(CandidatePairs.Count + 1, string.Empty, string.Empty, true));
        RefreshSelectedPairCount();
        await Task.CompletedTask;
    }

    [RelayCommand]
    private async Task RemovePairRowAsync(ProjectPairItemViewModel? item)
    {
        if (item is null)
        {
            return;
        }

        item.PropertyChanged -= PairItemPropertyChanged;
        CandidatePairs.Remove(item);
        RenumberPairs();
        RefreshSelectedPairCount();

        await Task.CompletedTask;
    }

    [RelayCommand]
    private async Task ApplySelectionAsync()
    {
        try
        {
            var selected = CandidatePairs.Where(x => x.IsSelected).ToList();
            if (selected.Count == 0)
            {
                selected = CandidatePairs.Where(x => !string.IsNullOrWhiteSpace(x.IValue) && !string.IsNullOrWhiteSpace(x.JValue)).ToList();
            }

            if (selected.Count == 0)
            {
                throw new InvalidOperationException("Select at least one project pair.");
            }

            EnsureDirectoryForFile(ActiveITablePath);
            EnsureDirectoryForFile(ActiveJTablePath);

            var iEncoding = ResolveTextEncodingForPath(ActiveITablePath);
            var jEncoding = ResolveTextEncodingForPath(ActiveJTablePath);

            await File.WriteAllLinesAsync(ActiveITablePath, selected.Select(x => NormalizeLineValue(x.IValue)), iEncoding);
            await File.WriteAllLinesAsync(ActiveJTablePath, selected.Select(x => NormalizeLineValue(x.JValue)), jEncoding);

            SelectedPairCount = selected.Count;
            StatusText = $"Saved {selected.Count} selected pairs.";
        }
        catch (Exception ex)
        {
            SetError($"Failed to save selection: {ex.Message}");
        }
    }

    [RelayCommand(CanExecute = nameof(CanStartTest))]
    private async Task StartTestAsync()
    {
        try
        {
            if (IsRunning)
            {
                return;
            }

            await ApplySelectionAsync();
            var launch = ResolveLaunchCommand();

            IsRunning = true;
            IsPaused = false;
            _stopRequestedByUser = false;
            _runStartedUtc = DateTime.UtcNow;
            RunningProcessId = 0;
            ConfiguredLoopCount = Math.Max(1, SelectedPairCount);
            LoopStartedCount = 0;
            LoopCompletedCount = 0;
            LoopFailedCount = 0;
            ErrorLineCount = 0;
            ProcessedSummaryCount = 0;
            TotalSuccessInSummaries = 0;
            TotalFailureInSummaries = 0;
            LastError = string.Empty;
            CurrentRunLogPath = Path.Combine(OutputLogDirectoryPath, $"gui_run_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            StatusText = "Starting...";
            RunResult = "RUNNING";
            VerdictText = "RUNNING";
            VerdictReason = "Test is running.";
            ProgressPercentage = 0;
            ProgressText = $"0 / {ConfiguredLoopCount} (0.0%)";
            EstimatedRemaining = "-";
            EstimatedFinishTime = "-";
            ElapsedTime = "00:00:00";
            OutputActivity = "active";
            LastOutputTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            LastOutputAgeSeconds = 0;

            Directory.CreateDirectory(OutputLogDirectoryPath);
            await File.WriteAllTextAsync(CurrentRunLogPath, "[GUI] Start requested\n", new UTF8Encoding(false));
            AppendLog("[GUI] Start requested.", false, true);
            await AppendRunLogLineAsync("[GUI] Start requested.");

            var process = BuildProcess(launch);
            _runningProcess = process;
            process.OutputDataReceived += (_, e) => HandleProcessOutputLine(e.Data, false);
            process.ErrorDataReceived += (_, e) => HandleProcessOutputLine(e.Data, true);

            if (!process.Start())
            {
                throw new InvalidOperationException("Failed to start command process.");
            }

            RunningProcessId = process.Id;
            StatusText = $"Running (PID {RunningProcessId})";
            AppendLog($"[GUI] Process started: {launch.DisplayCommand}", false, true);
            await AppendRunLogLineAsync($"[GUI] Process started: {launch.DisplayCommand}");

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            await process.WaitForExitAsync();

            try
            {
                process.CancelOutputRead();
                process.CancelErrorRead();
            }
            catch
            {
                // Ignore cancellation exceptions after process exit.
            }

            var exitCode = process.ExitCode;
            ElapsedTime = (DateTime.UtcNow - _runStartedUtc).ToString(@"hh\:mm\:ss");
            OutputActivity = "idle";

            if (_stopRequestedByUser)
            {
                RunResult = "STOPPED";
                VerdictText = "STOPPED";
                VerdictReason = "Stopped by user.";
                StatusText = "Stopped.";
                AppendLog($"[GUI] Process stopped by user (exit code {exitCode}).", false, true);
                await AppendRunLogLineAsync($"[GUI] Process stopped by user (exit code {exitCode}).");
                return;
            }

            if (exitCode == 0)
            {
                ProcessedSummaryCount = ConfiguredLoopCount;
                LoopStartedCount = ConfiguredLoopCount;
                LoopCompletedCount = ConfiguredLoopCount;
                TotalSuccessInSummaries = ConfiguredLoopCount;
                TotalFailureInSummaries = 0;
                ProgressPercentage = 100;
                ProgressText = $"{ConfiguredLoopCount} / {ConfiguredLoopCount} (100.0%)";
                RunResult = "PASS";
                VerdictText = "PASS";
                VerdictReason = "Process exited with code 0.";
                StatusText = "Completed.";
                AppendLog("[GUI] Completed.", false, true);
                await AppendRunLogLineAsync("[GUI] Completed.");
            }
            else
            {
                LoopFailedCount = 1;
                TotalFailureInSummaries = 1;
                RunResult = "FAIL";
                VerdictText = "FAIL";
                VerdictReason = $"Process exited with code {exitCode}.";
                StatusText = "Failed.";
                AppendLog($"[ERR] Process exited with code {exitCode}.", true, false);
                await AppendRunLogLineAsync($"[ERR] Process exited with code {exitCode}.");
            }
        }
        catch (Exception ex)
        {
            SetError($"Start failed: {ex.Message}");
            RunResult = "FAIL";
            VerdictText = "FAIL";
            VerdictReason = ex.Message;
        }
        finally
        {
            _runningProcess?.Dispose();
            _runningProcess = null;
            IsRunning = false;
            IsPaused = false;
            RunningProcessId = 0;
        }
    }

    [RelayCommand(CanExecute = nameof(CanStopTest))]
    private async Task StopTestAsync()
    {
        if (!IsRunning || _runningProcess is null)
        {
            StatusText = "No active process.";
            return;
        }

        if (_stopRequestedByUser)
        {
            return;
        }

        _stopRequestedByUser = true;
        StatusText = "Stopping...";
        RunResult = "STOPPING";
        VerdictText = "STOPPING";
        VerdictReason = "Stop requested by user.";
        AppendLog("[GUI] Stop requested by user.", false, true);
        await AppendRunLogLineAsync("[GUI] Stop requested by user.");

        await KillProcessTreeAsync(_runningProcess.Id);
    }

    [RelayCommand]
    private void ClearLogs()
    {
        LogLines.Clear();
        LiveOutputText = string.Empty;
    }

    private bool CanStartTest()
    {
        return !IsRunning;
    }

    private bool CanStopTest()
    {
        return IsRunning;
    }

    private async Task<string?> LoadEnvironmentFromSourceScriptAsync()
    {
        var path = ResolveEnvironmentScriptPath();
        if (path is null || !File.Exists(path))
        {
            return null;
        }

        var encoding = ResolveTextEncodingForPath(path);
        var lines = await File.ReadAllLinesAsync(path, encoding);
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in lines)
        {
            if (!TryParseEnvironmentAssignment(line, out var key, out var value))
            {
                continue;
            }

            if (IsTrackedEnvironmentKey(key))
            {
                values[key] = value;
            }
        }

        ApplyEnvironmentValues(values);
        AppendLog($"[GUI] Loaded environment values from {Path.GetFileName(path)}.", false, false);
        return path;
    }

    private void ApplyEnvironmentValues(Dictionary<string, string> values)
    {
        if (values.TryGetValue("QAF_ROOT", out var qafRoot))
        {
            QafRoot = qafRoot;
        }

        if (values.TryGetValue("QACLI_BIN", out var qacliBin))
        {
            QacliBinPath = qacliBin;
        }

        if (values.TryGetValue("TEST_ROOT", out var testRoot))
        {
            TestRoot = testRoot;
        }

        if (values.TryGetValue("COM_QAC", out var comQac))
        {
            ComQac = comQac;
        }

        if (values.TryGetValue("COM_QACPP", out var comQacpp))
        {
            ComQacpp = comQacpp;
        }

        if (values.TryGetValue("COM_RCMA", out var comRcma))
        {
            ComRcma = comRcma;
        }

        if (values.TryGetValue("COM_MTA", out var comMta))
        {
            ComMta = comMta;
        }

        if (values.TryGetValue("COM_NAME", out var comName))
        {
            ComName = comName;
        }

        if (values.TryGetValue("COM_DATA", out var comData))
        {
            ComData = comData;
        }

        if (values.TryGetValue("QAV_SERVER", out var qavServer))
        {
            QavServer = qavServer;
        }

        if (values.TryGetValue("QAV_USER", out var qavUser))
        {
            QavUser = qavUser;
        }

        if (values.TryGetValue("QAV_PASS", out var qavPass))
        {
            QavPass = qavPass;
        }

        if (values.TryGetValue("VAL_SERVER", out var valServer))
        {
            ValServer = valServer;
        }

        if (values.TryGetValue("VAL_USER", out var valUser))
        {
            ValUser = valUser;
        }

        if (values.TryGetValue("VAL_PASS", out var valPass))
        {
            ValPass = valPass;
        }
    }

    private string? ResolveEnvironmentScriptPath()
    {
        if (OperatingSystem.IsWindows())
        {
            return FindFirstExistingFile(LinURootPath, "test_loop.bat", "testloop.bat");
        }

        return FindFirstExistingFile(LinURootPath, "testonce.sh", "test_loop.bat", "testloop.bat");
    }

    private string ResolveEnvironmentScriptPathForSave()
    {
        if (OperatingSystem.IsWindows())
        {
            return FindFirstExistingFile(LinURootPath, "test_loop.bat", "testloop.bat")
                ?? Path.Combine(LinURootPath, "test_loop.bat");
        }

        var existing = ResolveEnvironmentScriptPath();
        if (existing is not null)
        {
            return existing;
        }

        return Path.Combine(LinURootPath, "testonce.sh");
    }

    private async Task SaveEnvironmentToShellScriptAsync(string path)
    {
        var lines = File.Exists(path)
            ? (await File.ReadAllLinesAsync(path, new UTF8Encoding(false))).ToList()
            : new List<string> { "#!/usr/bin/env bash", string.Empty };

        UpsertEnvironmentAssignments(lines, isBatch: false);
        await File.WriteAllLinesAsync(path, lines, new UTF8Encoding(false));
    }

    private async Task SaveEnvironmentToBatchScriptAsync(string path)
    {
        var encoding = Encoding.GetEncoding(932);
        var lines = File.Exists(path)
            ? (await File.ReadAllLinesAsync(path, encoding)).ToList()
            : new List<string> { "@ECHO OFF", string.Empty, "SETLOCAL ENABLEDELAYEDEXPANSION", string.Empty };

        UpsertEnvironmentAssignments(lines, isBatch: true);
        await File.WriteAllLinesAsync(path, lines, encoding);
    }

    private void UpsertEnvironmentAssignments(List<string> lines, bool isBatch)
    {
        foreach (var key in TrackedEnvironmentKeys)
        {
            var replacement = BuildEnvironmentLine(key, GetEnvironmentValue(key), isBatch);
            var indexes = new List<int>();
            for (var i = 0; i < lines.Count; i++)
            {
                if (TryParseEnvironmentAssignment(lines[i], out var lineKey, out _)
                    && string.Equals(lineKey, key, StringComparison.OrdinalIgnoreCase))
                {
                    indexes.Add(i);
                }
            }

            if (indexes.Count == 0)
            {
                lines.Add(replacement);
                continue;
            }

            lines[indexes[0]] = replacement;
            for (var i = indexes.Count - 1; i >= 1; i--)
            {
                lines.RemoveAt(indexes[i]);
            }
        }
    }

    private string GetEnvironmentValue(string key)
    {
        return key switch
        {
            "QAF_ROOT" => QafRoot,
            "QACLI_BIN" => QacliBinPath,
            "TEST_ROOT" => TestRoot,
            "COM_QAC" => ComQac,
            "COM_QACPP" => ComQacpp,
            "COM_RCMA" => ComRcma,
            "COM_MTA" => ComMta,
            "COM_NAME" => ComName,
            "COM_DATA" => ComData,
            "QAV_SERVER" => QavServer,
            "QAV_USER" => QavUser,
            "QAV_PASS" => QavPass,
            "VAL_SERVER" => ValServer,
            "VAL_USER" => ValUser,
            "VAL_PASS" => ValPass,
            _ => string.Empty
        };
    }

    private static string BuildEnvironmentLine(string key, string value, bool isBatch)
    {
        if (isBatch)
        {
            return $"SET {key}={value}";
        }

        var escaped = value.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);
        return $"export {key}=\"{escaped}\"";
    }

    private static bool TryParseEnvironmentAssignment(string line, out string key, out string value)
    {
        key = string.Empty;
        value = string.Empty;

        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var trimmed = line.Trim();
        if (trimmed.StartsWith("#", StringComparison.Ordinal)
            || trimmed.StartsWith("::", StringComparison.Ordinal)
            || trimmed.StartsWith("REM ", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (trimmed.StartsWith("SETLOCAL", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("ENDLOCAL", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (trimmed.StartsWith("export ", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseAssignmentBody(trimmed[7..].Trim(), out key, out value);
        }

        if (trimmed.StartsWith("set ", StringComparison.OrdinalIgnoreCase))
        {
            var body = trimmed[4..].TrimStart();
            if (body.StartsWith("\"", StringComparison.Ordinal) && body.EndsWith("\"", StringComparison.Ordinal) && body.Length >= 2)
            {
                body = body[1..^1];
            }

            return TryParseAssignmentBody(body, out key, out value);
        }

        return TryParseAssignmentBody(trimmed, out key, out value);
    }

    private static bool TryParseAssignmentBody(string body, out string key, out string value)
    {
        key = string.Empty;
        value = string.Empty;

        var eqIndex = body.IndexOf('=');
        if (eqIndex <= 0)
        {
            return false;
        }

        var keyCandidate = body[..eqIndex].Trim().Trim('"');
        if (!IsValidEnvironmentKey(keyCandidate))
        {
            return false;
        }

        var valueCandidate = body[(eqIndex + 1)..].Trim();
        if ((valueCandidate.StartsWith("\"", StringComparison.Ordinal) && valueCandidate.EndsWith("\"", StringComparison.Ordinal))
            || (valueCandidate.StartsWith("'", StringComparison.Ordinal) && valueCandidate.EndsWith("'", StringComparison.Ordinal)))
        {
            if (valueCandidate.Length >= 2)
            {
                valueCandidate = valueCandidate[1..^1];
            }
        }

        key = keyCandidate;
        value = valueCandidate;
        return true;
    }

    private static bool IsValidEnvironmentKey(string key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return false;
        }

        var first = key[0];
        if (!(char.IsLetter(first) || first == '_'))
        {
            return false;
        }

        for (var i = 1; i < key.Length; i++)
        {
            var c = key[i];
            if (!(char.IsLetterOrDigit(c) || c == '_'))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsTrackedEnvironmentKey(string key)
    {
        foreach (var target in TrackedEnvironmentKeys)
        {
            if (string.Equals(target, key, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private LaunchCommand ResolveLaunchCommand()
    {
        var windowsScript = FindFirstExistingFile(LinURootPath, "test_loop.bat", "testloop.bat");
        if (OperatingSystem.IsWindows() && windowsScript is not null)
        {
            return new LaunchCommand(
                "cmd.exe",
                $"/d /c \"\"{windowsScript}\"\"",
                LinURootPath,
                Encoding.GetEncoding(932),
                Encoding.GetEncoding(932),
                $"cmd.exe /d /c \"{windowsScript}\"");
        }

        var shellScript = FindFirstExistingFile(LinURootPath, "testloop.generated.sh", "testloop.sh", "test_loop.sh");
        if (shellScript is not null)
        {
            return new LaunchCommand(
                "bash",
                $"\"{shellScript}\"",
                LinURootPath,
                new UTF8Encoding(false),
                new UTF8Encoding(false),
                $"bash \"{shellScript}\"");
        }

        throw new FileNotFoundException("No runnable script found. Expected test_loop.bat, testloop.bat, or testloop.sh.");
    }

    private Process BuildProcess(LaunchCommand launch)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = launch.FileName,
            Arguments = launch.Arguments,
            WorkingDirectory = launch.WorkingDirectory,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = launch.StandardOutputEncoding,
            StandardErrorEncoding = launch.StandardErrorEncoding
        };

        SetEnvironmentIfNotEmpty(startInfo, "QAF_ROOT", QafRoot);
        SetEnvironmentIfNotEmpty(startInfo, "QACLI_BIN", QacliBinPath);
        SetEnvironmentIfNotEmpty(startInfo, "TEST_ROOT", TestRoot);
        SetEnvironmentIfNotEmpty(startInfo, "COM_QAC", ComQac);
        SetEnvironmentIfNotEmpty(startInfo, "COM_QACPP", ComQacpp);
        SetEnvironmentIfNotEmpty(startInfo, "COM_RCMA", ComRcma);
        SetEnvironmentIfNotEmpty(startInfo, "COM_MTA", ComMta);
        SetEnvironmentIfNotEmpty(startInfo, "COM_NAME", ComName);
        SetEnvironmentIfNotEmpty(startInfo, "COM_DATA", ComData);
        SetEnvironmentIfNotEmpty(startInfo, "QAV_SERVER", QavServer);
        SetEnvironmentIfNotEmpty(startInfo, "QAV_USER", QavUser);
        SetEnvironmentIfNotEmpty(startInfo, "QAV_PASS", QavPass);
        SetEnvironmentIfNotEmpty(startInfo, "VAL_SERVER", ValServer);
        SetEnvironmentIfNotEmpty(startInfo, "VAL_USER", ValUser);
        SetEnvironmentIfNotEmpty(startInfo, "VAL_PASS", ValPass);

        return new Process
        {
            StartInfo = startInfo,
            EnableRaisingEvents = true
        };
    }

    private void HandleProcessOutputLine(string? line, bool isError)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return;
        }

        Dispatcher.UIThread.Post(() =>
        {
            AppendLog(line, isError, false);
            LastOutputTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            LastOutputAgeSeconds = 0;
            OutputActivity = "active";

            if (isError || line.IndexOf("[ERR]", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                ErrorLineCount += 1;

                if (StopOnFirstError && !_stopRequestedByUser)
                {
                    _ = StopTestAsync();
                }
            }

            if (AutoConfirmPrompts && IsPausePrompt(line))
            {
                _ = SendAutoConfirmAsync();
            }

            _ = AppendRunLogLineAsync(line);
        });
    }

    private async Task SendAutoConfirmAsync()
    {
        var process = _runningProcess;
        if (process is null || process.HasExited)
        {
            return;
        }

        try
        {
            await process.StandardInput.WriteLineAsync(" ");
            await process.StandardInput.FlushAsync();

            AutoEnterCount += 1;
            AppendLog("[GUI] Auto-confirm sent.", false, true);
            await AppendRunLogLineAsync("[GUI] Auto-confirm sent.");
        }
        catch
        {
            // Ignore stdin write failures when process is terminating.
        }
    }

    private async Task AppendRunLogLineAsync(string text)
    {
        if (string.IsNullOrWhiteSpace(CurrentRunLogPath))
        {
            return;
        }

        await _logWriteSemaphore.WaitAsync();
        try
        {
            await File.AppendAllTextAsync(CurrentRunLogPath, text + Environment.NewLine, new UTF8Encoding(false));
        }
        catch
        {
            // Do not fail the run because of log write errors.
        }
        finally
        {
            _logWriteSemaphore.Release();
        }
    }

    private static bool IsPausePrompt(string line)
    {
        return (line.IndexOf("press any key", StringComparison.OrdinalIgnoreCase) >= 0)
            || line.IndexOf("Press [Enter]", StringComparison.OrdinalIgnoreCase) >= 0
            || line.IndexOf("続行するには何かキー", StringComparison.Ordinal) >= 0;
    }

    private static string? FindFirstExistingFile(string baseDirectory, params string[] fileNames)
    {
        foreach (var fileName in fileNames)
        {
            var path = Path.Combine(baseDirectory, fileName);
            if (File.Exists(path))
            {
                return path;
            }
        }

        return null;
    }

    private static void SetEnvironmentIfNotEmpty(ProcessStartInfo startInfo, string key, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        startInfo.Environment[key] = value;
    }

    private static async Task KillProcessTreeAsync(int pid)
    {
        try
        {
            if (OperatingSystem.IsWindows())
            {
                using var killer = Process.Start(new ProcessStartInfo
                {
                    FileName = "taskkill",
                    Arguments = $"/PID {pid} /T /F",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                });

                if (killer is not null)
                {
                    await killer.WaitForExitAsync();
                }

                return;
            }

            try
            {
                Process.GetProcessById(pid).Kill(entireProcessTree: true);
            }
            catch
            {
                // Ignore race where process has already exited.
            }
        }
        catch
        {
            // Ignore best-effort stop failures.
        }
    }

    private ProjectPairItemViewModel CreatePairItem(int index, string iValue, string jValue, bool isSelected)
    {
        var item = new ProjectPairItemViewModel(index, iValue, jValue, RemovePairRowInternal)
        {
            IsSelected = isSelected
        };

        item.PropertyChanged += PairItemPropertyChanged;
        return item;
    }

    private void RemovePairRowInternal(ProjectPairItemViewModel item)
    {
        _ = RemovePairRowAsync(item);
    }

    private void PairItemPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (string.Equals(e.PropertyName, nameof(ProjectPairItemViewModel.IsSelected), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(ProjectPairItemViewModel.IValue), StringComparison.Ordinal)
            || string.Equals(e.PropertyName, nameof(ProjectPairItemViewModel.JValue), StringComparison.Ordinal))
        {
            RefreshSelectedPairCount();
        }
    }

    private void RefreshSelectedPairCount()
    {
        SelectedPairCount = CandidatePairs.Count(x => x.IsSelected);
    }

    private void RenumberPairs()
    {
        for (var i = 0; i < CandidatePairs.Count; i++)
        {
            CandidatePairs[i].Index = i + 1;
        }
    }

    private void AppendLog(string text, bool isError, bool isCommand)
    {
        LogLines.Add(new LiveLogEntry(text, isError, isCommand));

        if (LogLines.Count > 4000)
        {
            LogLines.RemoveAt(0);
        }

        if (LiveOutputText.Length == 0)
        {
            LiveOutputText = text;
        }
        else
        {
            LiveOutputText += Environment.NewLine + text;
        }
    }

    private void SetError(string message)
    {
        LastError = message;
        StatusText = message;
        ErrorLineCount += 1;
        AppendLog($"[ERR] {message}", true, false);
        _ = AppendRunLogLineAsync($"[ERR] {message}");
    }

    private static string NormalizeLineValue(string input)
    {
        return input
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\uFEFF", string.Empty);
    }

    private static List<string> ReadLinesPreserveContent(string path)
    {
        if (!File.Exists(path))
        {
            return new List<string>();
        }

        var encoding = ResolveTextEncodingForPath(path);

        return File.ReadAllLines(path, encoding)
            .Select(x => x.Replace("\r", string.Empty).Replace("\uFEFF", string.Empty))
            .ToList();
    }

    private static Encoding ResolveTextEncodingForPath(string path)
    {
        if (path.EndsWith(".bat", StringComparison.OrdinalIgnoreCase)
            || path.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase)
            || path.IndexOf("SJIS", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return Encoding.GetEncoding(932);
        }

        return new UTF8Encoding(false);
    }

    private static void EnsureDirectoryForFile(string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dir))
        {
            Directory.CreateDirectory(dir);
        }
    }

    private static string DetectDefaultLinURoot()
    {
        var cwdCandidate = Path.Combine(Directory.GetCurrentDirectory(), "Lin_U");
        if (Directory.Exists(cwdCandidate))
        {
            return cwdCandidate;
        }

        var knownCandidate = "/home/itoke/Cusor_Project_Command_test/Lin_U";
        if (Directory.Exists(knownCandidate))
        {
            return knownCandidate;
        }

        return Directory.GetCurrentDirectory();
    }

    private static string DetectDefaultCommandTestDirectory()
    {
        var cwd = Directory.GetCurrentDirectory();

        if (Directory.Exists(Path.Combine(cwd, "Lin_U")))
        {
            return cwd;
        }

        if (string.Equals(Path.GetFileName(cwd), "Lin_U", StringComparison.OrdinalIgnoreCase))
        {
            return Directory.GetParent(cwd)?.FullName ?? cwd;
        }

        var knownCandidate = "/home/itoke/Cusor_Project_Command_test";
        if (Directory.Exists(knownCandidate))
        {
            return knownCandidate;
        }

        var linURoot = DetectDefaultLinURoot();
        return Directory.GetParent(linURoot)?.FullName ?? linURoot;
    }

    private sealed record LaunchCommand(
        string FileName,
        string Arguments,
        string WorkingDirectory,
        Encoding StandardOutputEncoding,
        Encoding StandardErrorEncoding,
        string DisplayCommand);
}

public partial class ProjectPairItemViewModel : ObservableObject
{
    public ProjectPairItemViewModel(int index, string iValue, string jValue, Action<ProjectPairItemViewModel>? removeAction = null)
    {
        Index = index;
        IValue = iValue;
        JValue = jValue;
        RemoveRowCommand = new RelayCommand(() =>
        {
            if (removeAction is not null)
            {
                removeAction(this);
            }
        });
    }

    [ObservableProperty]
    private int index;

    [ObservableProperty]
    private string iValue;

    [ObservableProperty]
    private string jValue;

    [ObservableProperty]
    private bool isSelected;

    public IRelayCommand RemoveRowCommand { get; }

    public static string MakeKey(string iValue, string jValue)
    {
        return iValue + "\u001F" + jValue;
    }
}

public sealed class LanguageOptionViewModel
{
    public LanguageOptionViewModel(string code, string displayName)
    {
        Code = code;
        DisplayName = displayName;
    }

    public string Code { get; }
    public string DisplayName { get; }
}

public sealed class LocalizationTexts : INotifyPropertyChanged
{
    private readonly Dictionary<string, Dictionary<string, string>> _texts = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ja"] = new(StringComparer.Ordinal)
        {
            ["HeaderTitle"] = "Lin_U QAC コマンドテストランナー (C# + Avalonia)",
            ["Language"] = "表示言語",
            ["SectionPaths"] = "パス設定",
            ["CommandTestDir"] = "コマンドテスト用ディレクトリ",
            ["BrowseDir"] = "ディレクトリ選択",
            ["LoadCommandDir"] = "指定ディレクトリを読込",
            ["LinURoot"] = "Lin_U ルート",
            ["AllIPath"] = "all I ファイル",
            ["AllJPath"] = "all J ファイル",
            ["ActiveIPath"] = "テスト I ファイル",
            ["ActiveJPath"] = "テスト J ファイル",
            ["ReloadPairs"] = "候補再読込",
            ["SelectFirst3"] = "先頭3件選択",
            ["ApplySelection"] = "選択を保存",
            ["SectionPairSelection"] = "プロジェクトペア選択 (I/J)",
            ["PairSelectionHint"] = "表示中の I/J は実際にテストされる内容です。編集・追加後に保存してください。",
            ["AddPairRow"] = "行を追加",
            ["DeleteRow"] = "削除",
            ["SelectedCount"] = "選択件数",
            ["SectionEnvironment"] = "環境上書き (環境スクリプト)",
            ["SaveEnvToScript"] = "GUI値を環境スクリプトへ保存",
            ["EnvLoadHint"] = "※ Windows は test_loop.bat、Linux は testonce.sh を自動読込します",
            ["QavCredential"] = "QAV サーバー / ユーザー / パスワード",
            ["ValCredential"] = "Validate サーバー / ユーザー / パスワード",
            ["SectionRunControl"] = "実行制御",
            ["StartTest"] = "テスト開始",
            ["Stop"] = "停止",
            ["ClearGuiLog"] = "GUIログクリア",
            ["AutoConfirm"] = "Press [Enter] を自動入力",
            ["StopOnError"] = "最初のエラーで停止",
            ["PromptDelay"] = "入力待ち遅延 (ms)",
            ["AutoEnterCount"] = "自動入力回数",
            ["Status"] = "ステータス",
            ["RunResult"] = "実行結果",
            ["Verdict"] = "判定",
            ["VerdictReason"] = "判定理由",
            ["ConfiguredLoops"] = "設定ループ数",
            ["RunningPid"] = "実行PID",
            ["Elapsed"] = "経過時間",
            ["Progress"] = "進捗",
            ["EstimatedFinish"] = "予想終了時刻",
            ["EstimatedRemaining"] = "残り時間",
            ["LoopStarted"] = "ループ開始数",
            ["LoopCompleted"] = "ループ完了数",
            ["LoopFailed"] = "ループ失敗数",
            ["ErrorLineCount"] = "ERR行数",
            ["OutputActivity"] = "出力状態",
            ["LastOutputTime"] = "最終出力時刻",
            ["LastOutputAge"] = "最終出力から",
            ["SummaryCount"] = "サマリ検出数",
            ["TotalSuccess"] = "サマリ成功合計",
            ["TotalFailure"] = "サマリ失敗合計",
            ["LastError"] = "最終エラー",
            ["RunLogFile"] = "実行ログファイル",
            ["SectionLiveOutput"] = "リアルタイム出力",
            ["AutoScrollLiveOutput"] = "自動スクロール",
            ["CopySelectedLog"] = "選択範囲をコピー"
        },
        ["en"] = new(StringComparer.Ordinal)
        {
            ["HeaderTitle"] = "Lin_U QAC Command Test Runner (C# + Avalonia)",
            ["Language"] = "Language",
            ["SectionPaths"] = "Paths",
            ["CommandTestDir"] = "Command Test Directory",
            ["BrowseDir"] = "Browse...",
            ["LoadCommandDir"] = "Load Directory",
            ["LinURoot"] = "Lin_U Root",
            ["AllIPath"] = "All I Path",
            ["AllJPath"] = "All J Path",
            ["ActiveIPath"] = "Test I File",
            ["ActiveJPath"] = "Test J File",
            ["ReloadPairs"] = "Reload Pairs",
            ["SelectFirst3"] = "Select First 3",
            ["ApplySelection"] = "Apply Selection",
            ["SectionPairSelection"] = "Project Pair Selection (I/J)",
            ["PairSelectionHint"] = "Displayed I/J pairs are used for actual tests. Edit/add and save.",
            ["AddPairRow"] = "Add Row",
            ["DeleteRow"] = "Delete",
            ["SelectedCount"] = "Selected Count",
            ["SectionEnvironment"] = "Environment Overrides (environment script)",
            ["SaveEnvToScript"] = "Save GUI values to environment script",
            ["EnvLoadHint"] = "* Windows loads test_loop.bat, Linux loads testonce.sh",
            ["QavCredential"] = "QAV Server / User / Password",
            ["ValCredential"] = "Validate Server / User / Password",
            ["SectionRunControl"] = "Run Control",
            ["StartTest"] = "Start Test",
            ["Stop"] = "Stop",
            ["ClearGuiLog"] = "Clear GUI Log",
            ["AutoConfirm"] = "Auto-confirm 'Press [Enter]'",
            ["StopOnError"] = "Stop on first error",
            ["PromptDelay"] = "Prompt delay (ms)",
            ["AutoEnterCount"] = "Auto-enter count",
            ["Status"] = "Status",
            ["RunResult"] = "Run Result",
            ["Verdict"] = "Verdict",
            ["VerdictReason"] = "Reason",
            ["ConfiguredLoops"] = "Configured Loops",
            ["RunningPid"] = "Process PID",
            ["Elapsed"] = "Elapsed",
            ["Progress"] = "Progress",
            ["EstimatedFinish"] = "Estimated Finish",
            ["EstimatedRemaining"] = "Remaining",
            ["LoopStarted"] = "Loops Started",
            ["LoopCompleted"] = "Loops Completed",
            ["LoopFailed"] = "Loops Failed",
            ["ErrorLineCount"] = "ERR Lines",
            ["OutputActivity"] = "Output Activity",
            ["LastOutputTime"] = "Last Output Time",
            ["LastOutputAge"] = "Output Age",
            ["SummaryCount"] = "Summary Count",
            ["TotalSuccess"] = "Total Success in Summaries",
            ["TotalFailure"] = "Total Failure in Summaries",
            ["LastError"] = "Last Error",
            ["RunLogFile"] = "Run Log File",
            ["SectionLiveOutput"] = "Live Output",
            ["AutoScrollLiveOutput"] = "Auto-scroll",
            ["CopySelectedLog"] = "Copy Selection"
        }
    };

    private string _languageCode = "ja";

    public event PropertyChangedEventHandler? PropertyChanged;

    public string this[string key]
    {
        get
        {
            if (_texts.TryGetValue(_languageCode, out var map) && map.TryGetValue(key, out var translated))
            {
                return translated;
            }

            if (_texts.TryGetValue("en", out var fallback) && fallback.TryGetValue(key, out var english))
            {
                return english;
            }

            return key;
        }
    }

    public void SetLanguage(string languageCode)
    {
        if (string.Equals(_languageCode, languageCode, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _languageCode = languageCode;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Item[]"));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Item"));
    }
}
