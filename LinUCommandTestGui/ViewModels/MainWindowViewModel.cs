using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
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
    public ObservableCollection<ParsedErrorItemViewModel> ErrorFindings { get; } = new();
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
    [NotifyCanExecuteChangedFor(nameof(PrecheckSettingsCommand))]
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
    private int extractedErrorCount;

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
    private bool _stopRequested;
    private bool _stopRequestedByUser;
    private bool _failureDetectedDuringRun;
    private string _firstFailureLine = string.Empty;
    private string _lastOutputLine = string.Empty;
    private string _lastErrorHintLine = string.Empty;
    private DateTime _lastOutputUtc = DateTime.UtcNow;
    private DateTime _lastAutoConfirmSentUtc = DateTime.MinValue;
    private int _observedOutputLineCount;
    private CancellationTokenSource? _autoConfirmLoopCts;
    private Task? _autoConfirmLoopTask;
    private CancellationTokenSource? _runtimeStatusLoopCts;
    private Task? _runtimeStatusLoopTask;
    private readonly Dictionary<string, ParsedErrorItemViewModel> _errorFindingIndex = new(StringComparer.Ordinal);
    private readonly SemaphoreSlim _logWriteSemaphore = new(1, 1);
    private static readonly HttpClient Http = new()
    {
        Timeout = TimeSpan.FromSeconds(5)
    };
    private static readonly Regex InvalidComponentRegex = new(
        @"(?:component name|コンポーネント名)\s*'(?<name>[^']+)'",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex JapaneseSummaryRegex = new(
        @"(?<success>\d+)\s*成功および、\s*(?<failed>\d+)\s*失敗",
        RegexOptions.Compiled);
    private static readonly Regex EnglishSummaryRegex = new(
        @"(?<success>\d+)\s+success(?:es)?\s*(?:and|,)\s*(?<failed>\d+)\s+fail(?:ed|ures?)?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GuiLoopMarkerRegex = new(
        @"^\[GUI-LOOP-(?<type>START|DONE|FAIL)\]\s*(?<index>\d+)?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
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
        StatusText = LocalizeText("\u5F85\u6A5F", "Idle");
        RunResult = LocalizeText("\u672A\u958B\u59CB", "Not started");
        VerdictText = LocalizeText("\u672A\u958B\u59CB", "Not started");
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
        OnPropertyChanged(nameof(PrecheckButtonText));
        OnPropertyChanged(nameof(ErrorAnalysisTabTitle));
        OnPropertyChanged(nameof(ExtractedErrorCountLabel));
        OnPropertyChanged(nameof(SecondUnitText));
        OnPropertyChanged(nameof(ErrorCategoryLabel));
        OnPropertyChanged(nameof(ErrorOccurrencesLabel));
        OnPropertyChanged(nameof(ErrorFirstSeenLabel));
        OnPropertyChanged(nameof(ErrorLastSeenLabel));
        OnPropertyChanged(nameof(ErrorExplanationLabel));
        OnPropertyChanged(nameof(ErrorHintLabel));
    }

    private bool IsJapaneseUiLanguage()
    {
        return string.Equals(SelectedLanguage?.Code, "ja", StringComparison.OrdinalIgnoreCase);
    }

    private string LocalizeText(string japanese, string english)
    {
        return IsJapaneseUiLanguage() ? japanese : english;
    }

    private string GetOutputActivityText(bool isActive)
    {
        return isActive
            ? LocalizeText("\u7A3C\u50CD\u4E2D", "active")
            : LocalizeText("\u5F85\u6A5F", "idle");
    }

    private string FormatRunningStatus(int pid)
    {
        return IsJapaneseUiLanguage()
            ? $"\u5B9F\u884C\u4E2D (PID {pid})"
            : $"Running (PID {pid})";
    }

    private string FormatRunningWithFailuresStatus(int pid)
    {
        return IsJapaneseUiLanguage()
            ? $"\u30A8\u30E9\u30FC\u691C\u51FA\u4E2D\u3067\u5B9F\u884C\u4E2D (PID {pid})"
            : $"Running with failures (PID {pid})";
    }

    public string PrecheckButtonText => LocalizeText("\u4E8B\u524D\u78BA\u8A8D", "Precheck");

    public string ErrorAnalysisTabTitle => LocalizeText("\u30A8\u30E9\u30FC\u89E3\u6790", "Error Analysis");

    public string ExtractedErrorCountLabel => LocalizeText("\u62BD\u51FA\u30A8\u30E9\u30FC\u6570", "Extracted Errors");

    public string SecondUnitText => LocalizeText("\u79D2", "s");

    public string ErrorCategoryLabel => LocalizeText("\u30AB\u30C6\u30B4\u30EA", "Category");

    public string ErrorOccurrencesLabel => LocalizeText("\u767A\u751F\u56DE\u6570", "Occurrences");

    public string ErrorFirstSeenLabel => LocalizeText("\u521D\u56DE\u691C\u51FA", "First Seen");

    public string ErrorLastSeenLabel => LocalizeText("\u6700\u7D42\u691C\u51FA", "Last Seen");

    public string ErrorExplanationLabel => LocalizeText("\u8AAC\u660E", "Explanation");

    public string ErrorHintLabel => LocalizeText("\u30D2\u30F3\u30C8", "Hint");

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
            var launch = ResolveLaunchCommand(forExecution: true);
            var configurationErrors = ValidateLaunchConfiguration(launch);
            if (configurationErrors.Count > 0)
            {
                throw new InvalidOperationException(BuildConfigurationErrorMessage(configurationErrors));
            }

            IsRunning = true;
            IsPaused = false;
            _stopRequested = false;
            _stopRequestedByUser = false;
            _failureDetectedDuringRun = false;
            _firstFailureLine = string.Empty;
            _lastOutputLine = string.Empty;
            _lastErrorHintLine = string.Empty;
            _lastOutputUtc = DateTime.UtcNow;
            _lastAutoConfirmSentUtc = DateTime.MinValue;
            _observedOutputLineCount = 0;
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
            ResetExtractedErrors();
            CurrentRunLogPath = Path.Combine(OutputLogDirectoryPath, $"gui_run_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            StatusText = LocalizeText("\u958B\u59CB\u4E2D...", "Starting...");
            RunResult = "RUNNING";
            VerdictText = "RUNNING";
            VerdictReason = LocalizeText("\u30C6\u30B9\u30C8\u5B9F\u884C\u4E2D\u3067\u3059\u3002", "Test is running.");
            ProgressPercentage = 0;
            ProgressText = $"0 / {ConfiguredLoopCount} (0.0%)";
            EstimatedRemaining = "-";
            EstimatedFinishTime = "-";
            ElapsedTime = "00:00:00";
            OutputActivity = GetOutputActivityText(isActive: true);
            LastOutputTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            LastOutputAgeSeconds = 0;

            Directory.CreateDirectory(OutputLogDirectoryPath);
            await File.WriteAllTextAsync(CurrentRunLogPath, string.Empty, new UTF8Encoding(false));
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
            StatusText = FormatRunningStatus(RunningProcessId);
            AppendLog($"[GUI] Process started: {launch.DisplayCommand}", false, true);
            await AppendRunLogLineAsync($"[GUI] Process started: {launch.DisplayCommand}");

            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            StartAutoConfirmLoop();
            StartRuntimeStatusLoop();
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
            OutputActivity = GetOutputActivityText(isActive: false);

            if (_stopRequestedByUser)
            {
                RunResult = "STOPPED";
                VerdictText = "STOPPED";
                VerdictReason = LocalizeText("\u30E6\u30FC\u30B6\u30FC\u306B\u3088\u308A\u505C\u6B62\u3055\u308C\u307E\u3057\u305F\u3002", "Stopped by user.");
                StatusText = LocalizeText("\u505C\u6B62\u3057\u307E\u3057\u305F\u3002", "Stopped.");
                AppendLog($"[GUI] Process stopped by user (exit code {exitCode}).", false, true);
                await AppendRunLogLineAsync($"[GUI] Process stopped by user (exit code {exitCode}).");
                return;
            }

            if (_stopRequested && StopOnFirstError && _failureDetectedDuringRun)
            {
                LoopFailedCount = Math.Max(LoopFailedCount, 1);
                TotalFailureInSummaries = Math.Max(TotalFailureInSummaries, 1);
                NormalizeSingleLoopCountersAfterExit();
                RunResult = "FAIL";
                VerdictText = "FAIL";
                var reason = BuildProcessFailureReason(
                    exitCode,
                    LocalizeText("\u6700\u521D\u306E\u30A8\u30E9\u30FC\u691C\u51FA\u306B\u3088\u308A\u505C\u6B62\u3057\u307E\u3057\u305F\u3002", "Stopped on first error."));
                VerdictReason = reason;
                LastError = reason;
                StatusText = LocalizeText("\u6700\u521D\u306E\u30A8\u30E9\u30FC\u3067\u505C\u6B62\u3057\u307E\u3057\u305F\u3002", "Stopped on first error.");
                AppendLog($"[ERR] {reason}", true, false);
                await AppendRunLogLineAsync($"[ERR] {reason}");
                return;
            }

            if (exitCode == 0 && !_failureDetectedDuringRun)
            {
                if (ProcessedSummaryCount == 0)
                {
                    ProcessedSummaryCount = ConfiguredLoopCount;
                }

                LoopStartedCount = Math.Max(LoopStartedCount, ConfiguredLoopCount);
                LoopCompletedCount = Math.Max(LoopCompletedCount, ConfiguredLoopCount);

                if (TotalSuccessInSummaries == 0 && TotalFailureInSummaries == 0)
                {
                    TotalSuccessInSummaries = ConfiguredLoopCount;
                }

                TotalFailureInSummaries = 0;
                ProgressPercentage = 100;
                UpdateProgressTextDisplay();
                UpdateEstimatedTimeDisplay(DateTime.UtcNow - _runStartedUtc);
                RunResult = "PASS";
                VerdictText = "PASS";
                VerdictReason = LocalizeText("\u51E6\u7406\u304C\u6B63\u5E38\u7D42\u4E86\u3057\u307E\u3057\u305F\u3002 (exit code 0)", "Process exited with code 0.");
                StatusText = LocalizeText("\u5B8C\u4E86\u3057\u307E\u3057\u305F\u3002", "Completed.");
                AppendLog("[GUI] Completed.", false, true);
                await AppendRunLogLineAsync("[GUI] Completed.");
            }
            else
            {
                LoopFailedCount = Math.Max(LoopFailedCount, 1);
                TotalFailureInSummaries = Math.Max(TotalFailureInSummaries, 1);
                NormalizeSingleLoopCountersAfterExit();
                RunResult = "FAIL";
                VerdictText = "FAIL";
                var reason = exitCode != 0
                    ? BuildProcessFailureReason(exitCode)
                    : BuildOutputFailureReason();
                VerdictReason = reason;
                LastError = reason;
                StatusText = exitCode != 0
                    ? LocalizeText("\u5931\u6557\u3057\u307E\u3057\u305F\u3002", "Failed.")
                    : LocalizeText("\u30A8\u30E9\u30FC\u691C\u51FA\u3067\u5B8C\u4E86\u3057\u307E\u3057\u305F\u3002", "Completed with errors.");
                AppendLog($"[ERR] {reason}", true, false);
                await AppendRunLogLineAsync($"[ERR] {reason}");
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
            await StopRuntimeStatusLoopAsync();
            await StopAutoConfirmLoopAsync();
            _runningProcess?.Dispose();
            _runningProcess = null;
            IsRunning = false;
            IsPaused = false;
            RunningProcessId = 0;
        }
    }

    [RelayCommand(CanExecute = nameof(CanPrecheckSettings))]
    private async Task PrecheckSettingsAsync()
    {
        try
        {
            var launch = ResolveLaunchCommand();
            var errors = ValidateLaunchConfiguration(launch);
            var warnings = new List<string>();

            await ValidateConnectivityAndCredentialsAsync(errors, warnings);

            if (errors.Count > 0)
            {
                var message = BuildConfigurationReport(errors, warnings);
                SetError(message);
                RunResult = "PRECHECK_FAIL";
                VerdictText = "FAIL";
                VerdictReason = LocalizeText("\u4E8B\u524D\u78BA\u8A8D\u3067\u30A8\u30E9\u30FC\u3092\u691C\u51FA\u3057\u307E\u3057\u305F\u3002", "Precheck failed.");
                return;
            }

            StatusText = warnings.Count == 0
                ? LocalizeText("\u4E8B\u524D\u78BA\u8A8D\u306F\u6B63\u5E38\u3067\u3059\u3002", "Precheck passed.")
                : LocalizeText("\u8B66\u544A\u3042\u308A\u3067\u4E8B\u524D\u78BA\u8A8D\u306F\u6B63\u5E38\u3067\u3059\u3002", "Precheck passed with warnings.");
            RunResult = "PRECHECK_OK";
            VerdictText = warnings.Count == 0 ? "PASS" : "WARN";
            VerdictReason = warnings.Count == 0
                ? LocalizeText("\u3059\u3079\u3066\u306E\u8A2D\u5B9A\u30C1\u30A7\u30C3\u30AF\u306B\u5408\u683C\u3057\u307E\u3057\u305F\u3002", "All configuration checks passed.")
                : BuildWarningsSummary(warnings);
            LastError = warnings.Count == 0 ? string.Empty : BuildWarningsSummary(warnings);
            AppendLog(
                warnings.Count == 0
                    ? "[GUI] Precheck passed."
                    : $"[GUI] Precheck passed with warnings: {BuildWarningsSummary(warnings)}",
                false,
                true);
            await AppendRunLogLineAsync(
                warnings.Count == 0
                    ? "[GUI] Precheck passed."
                    : $"[GUI] Precheck passed with warnings: {BuildWarningsSummary(warnings)}");
        }
        catch (Exception ex)
        {
            SetError($"Precheck failed: {ex.Message}");
            RunResult = "PRECHECK_FAIL";
            VerdictText = "FAIL";
            VerdictReason = ex.Message;
        }
    }

    [RelayCommand(CanExecute = nameof(CanStopTest))]
    private async Task StopTestAsync()
    {
        if (!IsRunning || _runningProcess is null)
        {
            StatusText = LocalizeText("\u5B9F\u884C\u4E2D\u30D7\u30ED\u30BB\u30B9\u306F\u3042\u308A\u307E\u305B\u3093\u3002", "No active process.");
            return;
        }

        if (_stopRequestedByUser)
        {
            return;
        }

        _stopRequested = true;
        _stopRequestedByUser = true;
        StatusText = LocalizeText("\u505C\u6B62\u4E2D...", "Stopping...");
        RunResult = "STOPPING";
        VerdictText = "STOPPING";
        VerdictReason = LocalizeText("\u30E6\u30FC\u30B6\u30FC\u304C\u505C\u6B62\u3092\u8981\u6C42\u3057\u307E\u3057\u305F\u3002", "Stop requested by user.");
        AppendLog("[GUI] Stop requested by user.", false, true);
        await AppendRunLogLineAsync("[GUI] Stop requested by user.");

        await KillProcessTreeAsync(_runningProcess.Id);
    }

    [RelayCommand]
    private void ClearLogs()
    {
        LogLines.Clear();
        LiveOutputText = string.Empty;
        ResetExtractedErrors();
    }

    private bool CanStartTest()
    {
        return !IsRunning;
    }

    private bool CanPrecheckSettings()
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

    private List<string> ValidateLaunchConfiguration(LaunchCommand launch)
    {
        var errors = new List<string>();

        if (!Directory.Exists(LinURootPath))
        {
            errors.Add($"Selected directory does not exist: {LinURootPath}");
        }

        ValidatePathField("TEST_ROOT", TestRoot, errors, required: true, expectDirectory: true);
        ValidatePathField("QAF_ROOT", QafRoot, errors, required: true, expectDirectory: true);
        ValidatePathField("QACLI_BIN", QacliBinPath, errors, required: true, expectDirectory: true);

        if (!string.IsNullOrWhiteSpace(QacliBinPath))
        {
            try
            {
                var resolvedQacliBin = ResolveConfiguredPath(QacliBinPath);
                var qacliFileName = OperatingSystem.IsWindows() ? "qacli.exe" : "qacli";
                var qacliPath = Path.Combine(resolvedQacliBin, qacliFileName);
                if (Directory.Exists(resolvedQacliBin) && !File.Exists(qacliPath))
                {
                    errors.Add($"QACLI_BIN does not contain {qacliFileName}: {resolvedQacliBin}");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"QACLI_BIN could not be resolved: {QacliBinPath} ({ex.Message})");
            }
        }

        if (launch.FileName.Equals("cmd.exe", StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrWhiteSpace(TestRoot))
        {
            try
            {
                var resolvedTestRoot = ResolveConfiguredPath(TestRoot);
                if (!Directory.Exists(resolvedTestRoot))
                {
                    errors.Add($"Windows TEST_ROOT not found: {resolvedTestRoot}");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Windows TEST_ROOT could not be resolved: {TestRoot} ({ex.Message})");
            }
        }

        return errors;
    }

    private async Task ValidateConnectivityAndCredentialsAsync(List<string> errors, List<string> warnings)
    {
        ValidateCredentialTriplet("QAV_SERVER", QavServer, "QAV_USER", QavUser, "QAV_PASS", QavPass, errors);
        ValidateCredentialTriplet("VAL_SERVER", ValServer, "VAL_USER", ValUser, "VAL_PASS", ValPass, errors);

        if (!string.IsNullOrWhiteSpace(QavServer))
        {
            await ValidateServerEndpointAsync("QAV_SERVER", QavServer, errors, warnings);
        }

        if (!string.IsNullOrWhiteSpace(ValServer))
        {
            await ValidateServerEndpointAsync("VAL_SERVER", ValServer, errors, warnings);
        }

        if (string.IsNullOrWhiteSpace(ValServer)
            || string.IsNullOrWhiteSpace(ValUser)
            || string.IsNullOrWhiteSpace(ValPass))
        {
            AddIssue(warnings, "VAL credential precheck skipped because VAL_SERVER/VAL_USER/VAL_PASS is incomplete.");
            return;
        }

        if (!TryResolveQacliExecutablePath(out var qacliPath, out var qacliError))
        {
            AddIssue(errors, qacliError);
            return;
        }

        await RunValidateAuthCheckAsync(qacliPath, errors, warnings);
    }

    private static void ValidateCredentialTriplet(
        string serverName,
        string serverValue,
        string userName,
        string userValue,
        string passName,
        string passValue,
        List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(serverValue))
        {
            AddIssue(errors, $"{serverName} is empty.");
        }

        if (string.IsNullOrWhiteSpace(userValue))
        {
            AddIssue(errors, $"{userName} is empty.");
        }

        if (string.IsNullOrWhiteSpace(passValue))
        {
            AddIssue(errors, $"{passName} is empty.");
        }
    }

    private async Task ValidateServerEndpointAsync(string fieldName, string rawUrl, List<string> errors, List<string> warnings)
    {
        if (!TryBuildServerUri(rawUrl, out var serverUri, out var reason))
        {
            AddIssue(errors, $"{fieldName} is not a valid http/https URL: {rawUrl} ({reason})");
            return;
        }

        try
        {
            _ = await Dns.GetHostAddressesAsync(serverUri.Host);
        }
        catch (Exception ex)
        {
            AddIssue(errors, $"{fieldName} host could not be resolved: {serverUri.Host} ({ex.Message})");
            return;
        }

        try
        {
            using var headRequest = new HttpRequestMessage(HttpMethod.Head, serverUri);
            using var headResponse = await Http.SendAsync(headRequest);
            if (headResponse.StatusCode == HttpStatusCode.MethodNotAllowed)
            {
                using var getRequest = new HttpRequestMessage(HttpMethod.Get, serverUri);
                using var _ = await Http.SendAsync(getRequest);
            }
        }
        catch (TaskCanceledException)
        {
            AddIssue(errors, $"{fieldName} connection timed out: {serverUri}");
        }
        catch (HttpRequestException ex)
        {
            AddIssue(errors, $"{fieldName} is not reachable: {serverUri} ({ex.Message})");
        }
        catch (Exception ex)
        {
            AddIssue(warnings, $"{fieldName} connectivity check returned an unexpected error: {ex.Message}");
        }
    }

    private static bool TryBuildServerUri(string rawUrl, out Uri serverUri, out string reason)
    {
        serverUri = default!;
        reason = string.Empty;

        if (string.IsNullOrWhiteSpace(rawUrl))
        {
            reason = "empty";
            return false;
        }

        var normalized = rawUrl.Trim();
        if (!Uri.TryCreate(normalized, UriKind.Absolute, out var uri))
        {
            reason = "malformed URL";
            return false;
        }

        if (!string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            reason = "scheme must be http or https";
            return false;
        }

        if (string.IsNullOrWhiteSpace(uri.Host))
        {
            reason = "host is empty";
            return false;
        }

        serverUri = uri;
        return true;
    }

    private bool TryResolveQacliExecutablePath(out string qacliPath, out string error)
    {
        qacliPath = string.Empty;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(QacliBinPath))
        {
            error = "QACLI_BIN is empty.";
            return false;
        }

        try
        {
            var resolvedQacliBin = ResolveConfiguredPath(QacliBinPath);
            var qacliFileName = OperatingSystem.IsWindows() ? "qacli.exe" : "qacli";
            var candidate = Path.Combine(resolvedQacliBin, qacliFileName);
            if (!File.Exists(candidate))
            {
                error = $"QACLI_BIN does not contain {qacliFileName}: {resolvedQacliBin}";
                return false;
            }

            qacliPath = candidate;
            return true;
        }
        catch (Exception ex)
        {
            error = $"QACLI_BIN could not be resolved: {QacliBinPath} ({ex.Message})";
            return false;
        }
    }

    private async Task RunValidateAuthCheckAsync(string qacliPath, List<string> errors, List<string> warnings)
    {
        var workingDirectory = Directory.Exists(LinURootPath)
            ? LinURootPath
            : Directory.GetCurrentDirectory();

        var arguments = $"auth --validate --username {QuoteArgument(ValUser)} --password {QuoteArgument(ValPass)} --url {QuoteArgument(ValServer)}";
        var result = await RunProcessForPrecheckAsync(qacliPath, arguments, workingDirectory, TimeSpan.FromSeconds(20));
        if (result.TimedOut)
        {
            AddIssue(errors, "VAL authentication precheck timed out (20s).");
            return;
        }

        if (result.ExitCode == 0)
        {
            return;
        }

        var detailLine = FirstNonEmptyLine(result.StdErr, result.StdOut);
        var hint = InferConfigurationHint(detailLine);
        var message = string.IsNullOrWhiteSpace(detailLine)
            ? $"VAL authentication failed. qacli exited with code {result.ExitCode}."
            : $"VAL authentication failed. {CompactLine(detailLine)}";
        if (!string.IsNullOrWhiteSpace(hint))
        {
            message += $" Hint: {hint}";
        }

        AddIssue(errors, message);

        if (string.IsNullOrWhiteSpace(detailLine))
        {
            AddIssue(warnings, "qacli auth --validate returned no detail output.");
        }
    }

    private static async Task<PrecheckProcessResult> RunProcessForPrecheckAsync(
        string fileName,
        string arguments,
        string workingDirectory,
        TimeSpan timeout)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
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
        if (!ReferenceEquals(completedTask, waitTask))
        {
            try
            {
                process.Kill(entireProcessTree: true);
            }
            catch
            {
                // Ignore best-effort timeout kill failures.
            }

            await waitTask;
            return new PrecheckProcessResult(
                process.ExitCode,
                await stdOutTask,
                await stdErrTask,
                true);
        }

        return new PrecheckProcessResult(
            process.ExitCode,
            await stdOutTask,
            await stdErrTask,
            false);
    }

    private static string QuoteArgument(string value)
    {
        var escaped = (value ?? string.Empty).Replace("\"", "\\\"", StringComparison.Ordinal);
        return $"\"{escaped}\"";
    }

    private static string FirstNonEmptyLine(string first, string second)
    {
        static string Find(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var parts = text
                .Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            return parts.FirstOrDefault() ?? string.Empty;
        }

        var primary = Find(first);
        return !string.IsNullOrWhiteSpace(primary) ? primary : Find(second);
    }

    private static string BuildConfigurationReport(List<string> errors, List<string> warnings)
    {
        var sb = new StringBuilder(BuildConfigurationErrorMessage(errors));
        if (warnings.Count > 0)
        {
            sb.Append(Environment.NewLine);
            sb.Append("Warnings:");
            foreach (var warning in warnings)
            {
                sb.Append(Environment.NewLine);
                sb.Append(" - ");
                sb.Append(warning);
            }
        }

        return sb.ToString();
    }

    private static string BuildWarningsSummary(List<string> warnings)
    {
        if (warnings.Count == 0)
        {
            return string.Empty;
        }

        var shown = warnings
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Take(3)
            .Select(CompactLine)
            .ToList();
        return shown.Count == 0
            ? string.Empty
            : string.Join(" | ", shown);
    }

    private static void AddIssue(List<string> target, string issue)
    {
        if (string.IsNullOrWhiteSpace(issue))
        {
            return;
        }

        if (!target.Contains(issue, StringComparer.Ordinal))
        {
            target.Add(issue);
        }
    }

    private void ValidatePathField(string fieldName, string configuredValue, List<string> errors, bool required, bool expectDirectory)
    {
        if (string.IsNullOrWhiteSpace(configuredValue))
        {
            if (required)
            {
                errors.Add($"{fieldName} is empty.");
            }

            return;
        }

        try
        {
            var resolved = ResolveConfiguredPath(configuredValue);
            if (expectDirectory && !Directory.Exists(resolved))
            {
                errors.Add($"{fieldName} directory not found: {resolved}");
            }
        }
        catch (Exception ex)
        {
            errors.Add($"{fieldName} could not be resolved: {configuredValue} ({ex.Message})");
        }
    }

    private string ResolveConfiguredPath(string configuredValue)
    {
        var value = configuredValue.Trim().Trim('"');
        if (!string.IsNullOrWhiteSpace(QafRoot))
        {
            value = value.Replace("%QAF_ROOT%", QafRoot, StringComparison.OrdinalIgnoreCase);
            value = value.Replace("$QAF_ROOT", QafRoot, StringComparison.Ordinal);
        }

        value = Environment.ExpandEnvironmentVariables(value);
        if (string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        if (value == "~")
        {
            value = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }
        else if (value.StartsWith("~/", StringComparison.Ordinal) || value.StartsWith("~\\", StringComparison.Ordinal))
        {
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            value = Path.Combine(home, value[2..]);
        }

        if (!Path.IsPathRooted(value))
        {
            value = Path.Combine(LinURootPath, value);
        }

        return Path.GetFullPath(value);
    }

    private static string BuildConfigurationErrorMessage(List<string> errors)
    {
        var sb = new StringBuilder();
        sb.Append("Configuration error(s):");
        foreach (var error in errors)
        {
            sb.Append(Environment.NewLine);
            sb.Append(" - ");
            sb.Append(error);
        }

        return sb.ToString();
    }

    private LaunchCommand ResolveLaunchCommand(bool forExecution = false)
    {
        var windowsScript = FindFirstExistingFile(LinURootPath, "test_loop.bat", "testloop.bat");
        if (OperatingSystem.IsWindows() && windowsScript is not null)
        {
            var launchScript = forExecution
                ? PrepareWindowsRuntimeBatchScript(windowsScript)
                : windowsScript;

            return new LaunchCommand(
                "cmd.exe",
                $"/d /c \"\"{launchScript}\"\"",
                LinURootPath,
                Encoding.GetEncoding(932),
                Encoding.GetEncoding(932),
                $"cmd.exe /d /c \"{launchScript}\"");
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

    private string PrepareWindowsRuntimeBatchScript(string sourceScriptPath)
    {
        var encoding = Encoding.GetEncoding(932);
        var sourceLines = File.ReadAllLines(sourceScriptPath, encoding);
        var runtimeLines = InstrumentWindowsLoopMarkers(sourceLines);
        var runtimeScriptPath = Path.Combine(
            RuntimeDirectoryPath,
            $"{Path.GetFileNameWithoutExtension(sourceScriptPath)}.gui.generated.bat");

        EnsureDirectoryForFile(runtimeScriptPath);
        File.WriteAllLines(runtimeScriptPath, runtimeLines, encoding);
        return runtimeScriptPath;
    }

    private static List<string> InstrumentWindowsLoopMarkers(IReadOnlyList<string> sourceLines)
    {
        if (sourceLines.Count == 0)
        {
            return [];
        }

        if (sourceLines.Any(line => line.IndexOf("[GUI-LOOP-START]", StringComparison.OrdinalIgnoreCase) >= 0))
        {
            return sourceLines.ToList();
        }

        var result = new List<string>(sourceLines.Count + 8);
        var injectedAtLoop = false;

        foreach (var line in sourceLines)
        {
            if (!injectedAtLoop && IsOuterLoopHeaderLine(line))
            {
                var indent = GetLeadingWhitespace(line) + "      ";
                result.Add("SET GUI_LOOP_INDEX=0");
                result.Add(line);
                result.Add($"{indent}SET /A GUI_LOOP_INDEX=!GUI_LOOP_INDEX!+1");
                result.Add($"{indent}ECHO [GUI-LOOP-START] !GUI_LOOP_INDEX!");
                injectedAtLoop = true;
                continue;
            }

            if (injectedAtLoop && IsOuterLoopIncrementLine(line))
            {
                var indent = GetLeadingWhitespace(line);
                result.Add($"{indent}ECHO [GUI-LOOP-DONE] !GUI_LOOP_INDEX!");
            }

            result.Add(line);
        }

        return result;
    }

    private static bool IsOuterLoopHeaderLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var trimmed = line.TrimStart();
        return trimmed.StartsWith("FOR /F", StringComparison.OrdinalIgnoreCase)
            && trimmed.IndexOf("KANJI_TABLE_I", StringComparison.OrdinalIgnoreCase) >= 0
            && trimmed.IndexOf("DO (", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool IsOuterLoopIncrementLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var trimmed = line.TrimStart();
        return trimmed.StartsWith("SET /A M=", StringComparison.OrdinalIgnoreCase)
            && trimmed.IndexOf("+1", StringComparison.Ordinal) >= 0;
    }

    private static string GetLeadingWhitespace(string line)
    {
        var index = 0;
        while (index < line.Length && char.IsWhiteSpace(line[index]))
        {
            index++;
        }

        return index == 0 ? string.Empty : line[..index];
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
            var normalizedLine = line.Trim();
            _observedOutputLineCount += 1;
            _lastOutputLine = normalizedLine;
            _lastOutputUtc = DateTime.UtcNow;
            AppendLog(line, isError, false);
            LastOutputTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            LastOutputAgeSeconds = 0;
            OutputActivity = GetOutputActivityText(isActive: true);
            UpdateRealtimeProgressFromOutput(normalizedLine);

            var failureDetected = IsFailureOutputLine(normalizedLine);
            if (failureDetected)
            {
                ErrorLineCount += 1;
                _lastErrorHintLine = normalizedLine;
                MarkFailureDetected(normalizedLine);
                RecordErrorFinding(normalizedLine);

                if (StopOnFirstError && !_stopRequested)
                {
                    _ = StopOnFirstErrorAsync(normalizedLine);
                }
            }
            else if (IsErrorHintOutputLine(normalizedLine))
            {
                _lastErrorHintLine = normalizedLine;
                RecordErrorFinding(normalizedLine);
            }

            if (AutoConfirmPrompts && IsPausePrompt(line))
            {
                _ = SendAutoConfirmAsync("prompt");
            }

            _ = AppendRunLogLineAsync(line);
        });
    }

    private void MarkFailureDetected(string line)
    {
        if (!_failureDetectedDuringRun)
        {
            _failureDetectedDuringRun = true;
            _firstFailureLine = line.Trim();
            RunResult = "FAIL";
            VerdictText = "FAIL";
            VerdictReason = IsJapaneseUiLanguage()
                ? $"\u51FA\u529B\u3067\u30A8\u30E9\u30FC\u3092\u691C\u51FA: {_firstFailureLine}"
                : $"Failure detected in output: {_firstFailureLine}";
            StatusText = RunningProcessId > 0
                ? FormatRunningWithFailuresStatus(RunningProcessId)
                : LocalizeText("\u30A8\u30E9\u30FC\u691C\u51FA\u4E2D\u3067\u5B9F\u884C\u4E2D", "Running with failures");
            AppendLog("[GUI] Failure detected in output. Verdict switched to FAIL.", true, true);
            _ = AppendRunLogLineAsync("[GUI] Failure detected in output. Verdict switched to FAIL.");
            return;
        }

        RunResult = "FAIL";
        VerdictText = "FAIL";
    }

    private async Task StopOnFirstErrorAsync(string reasonLine)
    {
        if (_stopRequested || _runningProcess is null)
        {
            return;
        }

        _stopRequested = true;
        RunResult = "STOPPING";
        VerdictText = "FAIL";
        VerdictReason = IsJapaneseUiLanguage()
            ? $"\u6700\u521D\u306E\u30A8\u30E9\u30FC\u3067\u505C\u6B62: {CompactLine(reasonLine)}"
            : $"Stopping on first error: {CompactLine(reasonLine)}";
        StatusText = LocalizeText("\u6700\u521D\u306E\u30A8\u30E9\u30FC\u3067\u505C\u6B62\u4E2D...", "Stopping on first error...");
        AppendLog("[GUI] Stop requested automatically due to first error.", true, true);
        await AppendRunLogLineAsync("[GUI] Stop requested automatically due to first error.");

        if (!_runningProcess.HasExited)
        {
            await KillProcessTreeAsync(_runningProcess.Id);
        }
    }

    private void StartAutoConfirmLoop()
    {
        _autoConfirmLoopCts?.Cancel();
        _autoConfirmLoopCts?.Dispose();
        _autoConfirmLoopCts = new CancellationTokenSource();
        _autoConfirmLoopTask = RunAutoConfirmLoopAsync(_autoConfirmLoopCts.Token);
    }

    private void StartRuntimeStatusLoop()
    {
        _runtimeStatusLoopCts?.Cancel();
        _runtimeStatusLoopCts?.Dispose();
        _runtimeStatusLoopCts = new CancellationTokenSource();
        _runtimeStatusLoopTask = RunRuntimeStatusLoopAsync(_runtimeStatusLoopCts.Token);
    }

    private async Task StopAutoConfirmLoopAsync()
    {
        if (_autoConfirmLoopCts is null)
        {
            return;
        }

        try
        {
            _autoConfirmLoopCts.Cancel();
            if (_autoConfirmLoopTask is not null)
            {
                await _autoConfirmLoopTask;
            }
        }
        catch (OperationCanceledException)
        {
            // Expected when cancellation happens while waiting delay.
        }
        finally
        {
            _autoConfirmLoopCts.Dispose();
            _autoConfirmLoopCts = null;
            _autoConfirmLoopTask = null;
        }
    }

    private async Task StopRuntimeStatusLoopAsync()
    {
        if (_runtimeStatusLoopCts is null)
        {
            return;
        }

        try
        {
            _runtimeStatusLoopCts.Cancel();
            if (_runtimeStatusLoopTask is not null)
            {
                await _runtimeStatusLoopTask;
            }
        }
        catch (OperationCanceledException)
        {
            // Expected when cancellation happens while waiting delay.
        }
        finally
        {
            _runtimeStatusLoopCts.Dispose();
            _runtimeStatusLoopCts = null;
            _runtimeStatusLoopTask = null;
        }
    }

    private async Task RunAutoConfirmLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(500, cancellationToken);

            if (!AutoConfirmPrompts)
            {
                continue;
            }

            var process = _runningProcess;
            if (process is null || process.HasExited)
            {
                return;
            }

            var cadenceMs = Math.Max(1000, PromptDelayMilliseconds);
            var now = DateTime.UtcNow;
            if ((now - _lastOutputUtc).TotalMilliseconds < cadenceMs)
            {
                continue;
            }

            if ((now - _lastAutoConfirmSentUtc).TotalMilliseconds < cadenceMs)
            {
                continue;
            }

            await SendAutoConfirmAsync("watchdog");
        }
    }

    private async Task RunRuntimeStatusLoopAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            await Task.Delay(1000, cancellationToken);

            if (!IsRunning)
            {
                continue;
            }

            var nowUtc = DateTime.UtcNow;
            var elapsed = nowUtc - _runStartedUtc;
            var outputAgeSeconds = Math.Max(0, (int)(nowUtc - _lastOutputUtc).TotalSeconds);
            var outputState = GetOutputActivityText(outputAgeSeconds <= 2);

            Dispatcher.UIThread.Post(() =>
            {
                if (!IsRunning)
                {
                    return;
                }

                ElapsedTime = elapsed.ToString(@"hh\:mm\:ss");
                LastOutputAgeSeconds = outputAgeSeconds;
                OutputActivity = outputState;
                UpdateEstimatedTimeDisplay(elapsed);
            });
        }
    }

    private async Task SendAutoConfirmAsync(string reason = "prompt")
    {
        var process = _runningProcess;
        if (process is null || process.HasExited)
        {
            return;
        }

        try
        {
            await process.StandardInput.WriteLineAsync(string.Empty);
            await process.StandardInput.FlushAsync();
            _lastAutoConfirmSentUtc = DateTime.UtcNow;

            AutoEnterCount += 1;
            if (!string.Equals(reason, "watchdog", StringComparison.Ordinal))
            {
                const string message = "[GUI] Auto-confirm sent.";
                AppendLog(message, false, true);
                await AppendRunLogLineAsync(message);
            }
        }
        catch
        {
            // Ignore stdin write failures when process is terminating.
        }
    }

    private async Task AppendRunLogLineAsync(string text)
    {
        if (IsInternalGuiTraceLine(text))
        {
            return;
        }

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
        if (line.IndexOf("press any key", StringComparison.OrdinalIgnoreCase) >= 0
            || line.IndexOf("Press [Enter]", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return true;
        }

        if (line.IndexOf("続行するには何かキー", StringComparison.Ordinal) >= 0
            || line.IndexOf("何かキーを押してください", StringComparison.Ordinal) >= 0
            || line.IndexOf("邯夊｡後☆繧九↓縺ｯ菴輔°繧ｭ繝ｼ", StringComparison.Ordinal) >= 0)
        {
            return true;
        }

        return line.IndexOf(". . .", StringComparison.Ordinal) >= 0
            && (line.IndexOf("キー", StringComparison.Ordinal) >= 0
                || line.IndexOf("key", StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private void NormalizeSingleLoopCountersAfterExit()
    {
        if (ConfiguredLoopCount != 1)
        {
            return;
        }

        if (LoopStartedCount == 0)
        {
            LoopStartedCount = 1;
        }

        if (LoopCompletedCount == 0)
        {
            LoopCompletedCount = 1;
        }

        if (ProcessedSummaryCount == 0)
        {
            ProcessedSummaryCount = 1;
        }

        ProgressPercentage = 100;
        UpdateProgressTextDisplay();
    }

    private void UpdateRealtimeProgressFromOutput(string line)
    {
        if (!IsRunning || ConfiguredLoopCount <= 0)
        {
            return;
        }

        if (LoopStartedCount == 0)
        {
            LoopStartedCount = 1;
        }

        TryApplyGuiLoopMarker(line);

        if (TryParseSummaryCounts(line, out var successCount, out var failureCount))
        {
            ProcessedSummaryCount += 1;
            TotalSuccessInSummaries += successCount;
            TotalFailureInSummaries += failureCount;
        }

        var pulseProgress = Math.Min(95.0, 95.0 * (1.0 - Math.Exp(-_observedOutputLineCount / 280.0)));
        var summaryProgress = 0.0;
        if (ProcessedSummaryCount > 0)
        {
            const int estimatedSummariesPerLoop = 12;
            var estimatedTotalSummaries = Math.Max(1, ConfiguredLoopCount * estimatedSummariesPerLoop);
            summaryProgress = Math.Min(95.0, (ProcessedSummaryCount * 100.0) / estimatedTotalSummaries);
        }

        var nextProgress = Math.Max(ProgressPercentage, Math.Max(pulseProgress, summaryProgress));
        if (nextProgress > ProgressPercentage)
        {
            ProgressPercentage = nextProgress;
        }

        UpdateProgressTextDisplay();
        UpdateEstimatedTimeDisplay(DateTime.UtcNow - _runStartedUtc);
    }

    private void UpdateProgressTextDisplay()
    {
        if (ConfiguredLoopCount <= 0)
        {
            ProgressText = "0 / 0 (0.0%)";
            return;
        }

        var completed = Math.Min(LoopCompletedCount, ConfiguredLoopCount);
        ProgressText = $"{completed} / {ConfiguredLoopCount} ({ProgressPercentage:0.0}%)";
    }

    private void UpdateEstimatedTimeDisplay(TimeSpan elapsed)
    {
        var progress = Math.Clamp(ProgressPercentage, 0.0, 100.0);
        if (progress <= 0.01)
        {
            EstimatedRemaining = "-";
            EstimatedFinishTime = "-";
            return;
        }

        if (progress >= 99.9)
        {
            EstimatedRemaining = "00:00:00";
            EstimatedFinishTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            return;
        }

        var elapsedSeconds = Math.Max(1.0, elapsed.TotalSeconds);
        var estimatedTotalSeconds = elapsedSeconds * (100.0 / progress);
        var remainingSeconds = Math.Max(0.0, estimatedTotalSeconds - elapsedSeconds);
        if (double.IsNaN(remainingSeconds) || double.IsInfinity(remainingSeconds) || remainingSeconds > TimeSpan.FromDays(30).TotalSeconds)
        {
            EstimatedRemaining = "-";
            EstimatedFinishTime = "-";
            return;
        }

        var remaining = TimeSpan.FromSeconds(remainingSeconds);
        EstimatedRemaining = FormatDurationForDisplay(remaining);
        EstimatedFinishTime = DateTime.Now.Add(remaining).ToString("yyyy-MM-dd HH:mm:ss");
    }

    private static string FormatDurationForDisplay(TimeSpan duration)
    {
        if (duration < TimeSpan.Zero)
        {
            duration = TimeSpan.Zero;
        }

        var totalHours = (int)Math.Floor(duration.TotalHours);
        return $"{totalHours:00}:{duration.Minutes:00}:{duration.Seconds:00}";
    }

    private void TryApplyGuiLoopMarker(string line)
    {
        var match = GuiLoopMarkerRegex.Match(line);
        if (!match.Success)
        {
            return;
        }

        var markerType = match.Groups["type"].Value.ToUpperInvariant();
        var hasIndex = int.TryParse(match.Groups["index"].Value, out var parsedIndex) && parsedIndex > 0;
        var markerIndex = hasIndex ? parsedIndex : 0;

        switch (markerType)
        {
            case "START":
            {
                var started = hasIndex ? markerIndex : LoopStartedCount + 1;
                LoopStartedCount = Math.Min(ConfiguredLoopCount, Math.Max(LoopStartedCount, started));
                break;
            }
            case "DONE":
            {
                var completed = hasIndex ? markerIndex : LoopCompletedCount + 1;
                LoopCompletedCount = Math.Min(ConfiguredLoopCount, Math.Max(LoopCompletedCount, completed));
                LoopStartedCount = Math.Max(LoopStartedCount, LoopCompletedCount);
                break;
            }
            case "FAIL":
            {
                var failed = hasIndex ? markerIndex : LoopFailedCount + 1;
                LoopFailedCount = Math.Min(ConfiguredLoopCount, Math.Max(LoopFailedCount, failed));
                LoopStartedCount = Math.Max(LoopStartedCount, LoopFailedCount);
                break;
            }
        }

        var markerProgress = (Math.Min(LoopCompletedCount, ConfiguredLoopCount) * 100.0) / ConfiguredLoopCount;
        if (markerProgress > ProgressPercentage)
        {
            ProgressPercentage = markerProgress;
        }
    }

    private static bool TryParseSummaryCounts(string line, out int successCount, out int failureCount)
    {
        successCount = 0;
        failureCount = 0;
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var normalized = line.Trim();
        var jaMatch = JapaneseSummaryRegex.Match(normalized);
        if (jaMatch.Success
            && int.TryParse(jaMatch.Groups["success"].Value, out successCount)
            && int.TryParse(jaMatch.Groups["failed"].Value, out failureCount))
        {
            return true;
        }

        var enMatch = EnglishSummaryRegex.Match(normalized);
        if (enMatch.Success
            && int.TryParse(enMatch.Groups["success"].Value, out successCount)
            && int.TryParse(enMatch.Groups["failed"].Value, out failureCount))
        {
            return true;
        }

        return false;
    }

    private static bool IsFailureOutputLine(string line)
    {
        if (IsNonErrorSummaryLine(line) || IsCompatibilityLimitationLine(line))
        {
            return false;
        }

        if (IsLikelyPathListingLine(line))
        {
            return false;
        }

        if (line.IndexOf("[ERR]", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return true;
        }

        var lower = line.ToLowerInvariant();
        if (lower.Contains("not recognized as an internal or external command", StringComparison.Ordinal)
            || lower.Contains("the system cannot find", StringComparison.Ordinal)
            || lower.Contains("command not found", StringComparison.Ordinal)
            || lower.Contains("no such file or directory", StringComparison.Ordinal)
            || lower.Contains("fatal", StringComparison.Ordinal)
            || lower.Contains("unhandled exception", StringComparison.Ordinal)
            || lower.StartsWith("exception:", StringComparison.Ordinal)
            || lower.StartsWith("exception ", StringComparison.Ordinal))
        {
            return true;
        }

        if (line.IndexOf("見つかりません", StringComparison.Ordinal) >= 0
            || line.IndexOf("エラー", StringComparison.Ordinal) >= 0
            || line.IndexOf("失敗", StringComparison.Ordinal) >= 0)
        {
            return true;
        }

        if (ContainsFailureWord(lower, "error")
            || ContainsFailureWord(lower, "failed")
            || ContainsFailureWord(lower, "failure"))
        {
            return true;
        }

        return false;
    }

    private static bool IsBenignStdErrLine(string line)
    {
        if (IsLikelyPathListingLine(line))
        {
            return true;
        }

        var lower = line.ToLowerInvariant();
        return lower.Contains("files copied", StringComparison.Ordinal)
            || lower.Contains("directories moved", StringComparison.Ordinal);
    }

    private static bool IsLikelyPathListingLine(string line)
    {
        var trimmed = line.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return false;
        }

        if (!(trimmed.Contains('\\') || trimmed.Contains('/')))
        {
            return false;
        }

        if (trimmed.Contains(' '))
        {
            return false;
        }

        if (trimmed.StartsWith("[", StringComparison.Ordinal))
        {
            return false;
        }

        return true;
    }

    private static bool IsErrorHintOutputLine(string line)
    {
        if (IsNonErrorSummaryLine(line) || IsCompatibilityLimitationLine(line))
        {
            return false;
        }

        var lower = line.ToLowerInvariant();
        return lower.Contains("unauthorized", StringComparison.Ordinal)
            || lower.Contains("forbidden", StringComparison.Ordinal)
            || lower.Contains("authentication failed", StringComparison.Ordinal)
            || lower.Contains("invalid username", StringComparison.Ordinal)
            || lower.Contains("invalid user", StringComparison.Ordinal)
            || lower.Contains("invalid password", StringComparison.Ordinal)
            || lower.Contains("login failed", StringComparison.Ordinal)
            || lower.Contains("connection refused", StringComparison.Ordinal)
            || lower.Contains("actively refused", StringComparison.Ordinal)
            || lower.Contains("timed out", StringComparison.Ordinal)
            || lower.Contains("could not resolve", StringComparison.Ordinal)
            || lower.Contains("name or service not known", StringComparison.Ordinal)
            || lower.Contains("no such host", StringComparison.Ordinal);
    }

    private string BuildProcessFailureReason(int exitCode, string? prefix = null)
    {
        var line = GetMostRelevantFailureLine();
        var hint = InferConfigurationHint(line);
        var sb = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(prefix))
        {
            sb.Append(prefix.Trim());
            sb.Append(' ');
        }

        if (IsJapaneseUiLanguage())
        {
            sb.Append($"\u30D7\u30ED\u30BB\u30B9\u306F exit code {exitCode} \u3067\u7D42\u4E86\u3057\u307E\u3057\u305F\u3002");
        }
        else
        {
            sb.Append($"Process exited with code {exitCode}.");
        }

        if (!string.IsNullOrWhiteSpace(line))
        {
            sb.Append(IsJapaneseUiLanguage() ? " \u6700\u7D42\u30A8\u30E9\u30FC\u884C: " : " Last error line: ");
            sb.Append(CompactLine(line));
        }

        if (!string.IsNullOrWhiteSpace(hint))
        {
            sb.Append(IsJapaneseUiLanguage() ? " \u30D2\u30F3\u30C8: " : " Hint: ");
            sb.Append(hint);
        }

        return sb.ToString().Trim();
    }

    private string BuildOutputFailureReason()
    {
        var line = GetMostRelevantFailureLine();
        var hint = InferConfigurationHint(line);
        var sb = new StringBuilder(
            IsJapaneseUiLanguage()
                ? "\u51FA\u529B\u304B\u3089\u30A8\u30E9\u30FC\u3092\u691C\u51FA\u3057\u307E\u3057\u305F\u3002"
                : "Failure detected in output.");
        if (!string.IsNullOrWhiteSpace(line))
        {
            sb.Append(IsJapaneseUiLanguage() ? " \u6700\u7D42\u30A8\u30E9\u30FC\u884C: " : " Last error line: ");
            sb.Append(CompactLine(line));
        }

        if (!string.IsNullOrWhiteSpace(hint))
        {
            sb.Append(IsJapaneseUiLanguage() ? " \u30D2\u30F3\u30C8: " : " Hint: ");
            sb.Append(hint);
        }

        return sb.ToString();
    }

    private string GetMostRelevantFailureLine()
    {
        if (!string.IsNullOrWhiteSpace(_firstFailureLine))
        {
            return _firstFailureLine;
        }

        if (!string.IsNullOrWhiteSpace(_lastErrorHintLine))
        {
            return _lastErrorHintLine;
        }

        if (!string.IsNullOrWhiteSpace(_lastOutputLine))
        {
            return _lastOutputLine;
        }

        return string.Empty;
    }

    private string InferConfigurationHint(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return string.Empty;
        }

        var lower = line.ToLowerInvariant();

        if (lower.Contains("unauthorized", StringComparison.Ordinal)
            || lower.Contains("forbidden", StringComparison.Ordinal)
            || lower.Contains("authentication failed", StringComparison.Ordinal)
            || lower.Contains("invalid username", StringComparison.Ordinal)
            || lower.Contains("invalid user", StringComparison.Ordinal)
            || lower.Contains("invalid password", StringComparison.Ordinal)
            || lower.Contains("login failed", StringComparison.Ordinal)
            || lower.Contains("401", StringComparison.Ordinal))
        {
            return LocalizeText(
                "QAV_USER/QAV_PASS \u304A\u3088\u3073 VAL_USER/VAL_PASS \u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                "Check QAV_USER/QAV_PASS and VAL_USER/VAL_PASS.");
        }

        if (lower.Contains("could not resolve", StringComparison.Ordinal)
            || lower.Contains("name or service not known", StringComparison.Ordinal)
            || lower.Contains("no such host", StringComparison.Ordinal)
            || lower.Contains("host not found", StringComparison.Ordinal))
        {
            return LocalizeText(
                "QAV_SERVER / VAL_SERVER \u306E\u30DB\u30B9\u30C8\u540D\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                "Check QAV_SERVER / VAL_SERVER host names.");
        }

        if (lower.Contains("connection refused", StringComparison.Ordinal)
            || lower.Contains("actively refused", StringComparison.Ordinal)
            || lower.Contains("timed out", StringComparison.Ordinal)
            || lower.Contains("timeout", StringComparison.Ordinal)
            || lower.Contains("failed to connect", StringComparison.Ordinal))
        {
            return LocalizeText(
                "\u30B5\u30FC\u30D0\u30FC\u306E\u30A2\u30C9\u30EC\u30B9/\u30DD\u30FC\u30C8\u3068\u30B5\u30FC\u30D0\u30FC\u8D77\u52D5\u72B6\u614B\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                "Check server address/port and whether the server is running.");
        }

        if (lower.Contains("qacli is not recognized", StringComparison.Ordinal)
            || lower.Contains("qacli: command not found", StringComparison.Ordinal)
            || lower.Contains("not recognized as an internal or external command", StringComparison.Ordinal))
        {
            return LocalizeText(
                "QACLI_BIN \u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002qacli \u5B9F\u884C\u30D5\u30A1\u30A4\u30EB\u304C\u898B\u3064\u304B\u3089\u306A\u3044\u53EF\u80FD\u6027\u304C\u3042\u308A\u307E\u3059\u3002",
                "Check QACLI_BIN. qacli executable may not be found.");
        }

        if (lower.Contains("the system cannot find the path specified", StringComparison.Ordinal)
            || lower.Contains("path not found", StringComparison.Ordinal)
            || lower.Contains("no such file or directory", StringComparison.Ordinal))
        {
            return LocalizeText(
                "TEST_ROOT / QAF_ROOT / QACLI_BIN \u306E\u30D1\u30B9\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002",
                "Check TEST_ROOT / QAF_ROOT / QACLI_BIN paths.");
        }

        return string.Empty;
    }

    private static string CompactLine(string line)
    {
        var trimmed = line.Trim();
        return trimmed.Length <= 220 ? trimmed : trimmed[..220] + "...";
    }

    private static bool ContainsFailureWord(string lowerLine, string keyword)
    {
        var start = 0;
        while (true)
        {
            var index = lowerLine.IndexOf(keyword, start, StringComparison.Ordinal);
            if (index < 0)
            {
                return false;
            }

            var before = index > 0 ? lowerLine[index - 1] : ' ';
            var afterIndex = index + keyword.Length;
            var after = afterIndex < lowerLine.Length ? lowerLine[afterIndex] : ' ';
            var hasWordBoundary = !IsWordChar(before) && !IsWordChar(after);

            if (hasWordBoundary)
            {
                if (!(lowerLine.Contains($"0 {keyword}", StringComparison.Ordinal)
                    || lowerLine.Contains($"no {keyword}", StringComparison.Ordinal)
                    || lowerLine.Contains($"without {keyword}", StringComparison.Ordinal)))
                {
                    return true;
                }
            }

            start = index + keyword.Length;
        }
    }

    private static bool IsWordChar(char c)
    {
        return char.IsLetterOrDigit(c) || c == '_' || c == '-';
    }

    private static bool IsCompatibilityLimitationLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var normalized = line.Trim();
        var lower = normalized.ToLowerInvariant();
        if (IsQacliViewOptionParseLine(normalized, lower))
        {
            return true;
        }

        var componentMatch = InvalidComponentRegex.Match(normalized);
        return componentMatch.Success
            && (normalized.Contains("無効", StringComparison.Ordinal)
                || normalized.Contains("辟｡蜉ｹ", StringComparison.Ordinal)
                || lower.Contains("invalid", StringComparison.Ordinal));
    }

    private static bool IsQacliViewOptionParseLine(string normalized, string lower)
    {
        var hasParseMarker = normalized.Contains("パースエラー", StringComparison.Ordinal)
            || normalized.Contains("繝代・繧ｹ繧ｨ繝ｩ繝ｼ", StringComparison.Ordinal)
            || lower.Contains("parse error", StringComparison.Ordinal);
        if (!hasParseMarker)
        {
            return false;
        }

        return normalized.Contains("-t (--type)", StringComparison.Ordinal)
            || normalized.Contains("-m (--medium)", StringComparison.Ordinal)
            || lower.Contains("--type", StringComparison.Ordinal)
            || lower.Contains("--medium", StringComparison.Ordinal)
            || lower.Contains("argument:-t", StringComparison.Ordinal)
            || lower.Contains("argument:-m", StringComparison.Ordinal);
    }

    private static bool IsNonErrorSummaryLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var normalized = line.Trim();
        var lower = normalized.ToLowerInvariant();

        if (lower.Contains("0 errors", StringComparison.Ordinal)
            || lower.Contains("errors=0", StringComparison.Ordinal)
            || lower.Contains("0 failed", StringComparison.Ordinal)
            || lower.Contains("failures=0", StringComparison.Ordinal)
            || lower.Contains("ok, 0 error", StringComparison.Ordinal)
            || lower.Contains("0 error", StringComparison.Ordinal)
            || lower.Contains("no errors", StringComparison.Ordinal))
        {
            return true;
        }

        if (normalized.Contains("エラー=0", StringComparison.Ordinal)
            || normalized.Contains("エラー 0", StringComparison.Ordinal)
            || normalized.Contains("0 エラー", StringComparison.Ordinal)
            || normalized.Contains("失敗=0", StringComparison.Ordinal)
            || normalized.Contains("失敗 0", StringComparison.Ordinal)
            || normalized.Contains("0 失敗", StringComparison.Ordinal)
            || normalized.Contains("成功および、0 失敗", StringComparison.Ordinal)
            || normalized.Contains("OK, 0 エラー", StringComparison.Ordinal)
            || normalized.Contains("エラー 0 無効", StringComparison.Ordinal))
        {
            return true;
        }

        return false;
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
        if (IsInternalGuiTraceLine(text))
        {
            return;
        }

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

    private static bool IsInternalGuiTraceLine(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        return text.TrimStart().StartsWith("[GUI]", StringComparison.OrdinalIgnoreCase);
    }

    private void ResetExtractedErrors()
    {
        ErrorFindings.Clear();
        _errorFindingIndex.Clear();
        ExtractedErrorCount = 0;
    }

    private void RecordErrorFinding(string line)
    {
        if (!TryDescribeError(line, out var insight))
        {
            return;
        }

        var normalized = CompactLine(line.Trim());
        var key = $"{insight.Category}\u001F{normalized}";
        if (_errorFindingIndex.TryGetValue(key, out var existing))
        {
            existing.MarkOccurrence();
            RefreshExtractedErrorCount();
            return;
        }

        var item = new ParsedErrorItemViewModel(
            ErrorFindings.Count + 1,
            insight.Category,
            normalized,
            insight.Explanation,
            insight.Hint);
        ErrorFindings.Add(item);
        _errorFindingIndex[key] = item;
        RefreshExtractedErrorCount();
    }

    private void RefreshExtractedErrorCount()
    {
        ExtractedErrorCount = ErrorFindings.Sum(x => Math.Max(1, x.Occurrences));
    }

    private bool TryDescribeError(string line, out ErrorInsight insight)
    {
        insight = default!;
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        var normalized = line.Trim();
        if (IsNonErrorSummaryLine(normalized))
        {
            return false;
        }

        if (IsCompatibilityLimitationLine(normalized))
        {
            return false;
        }

        var componentMatch = InvalidComponentRegex.Match(normalized);
        if (componentMatch.Success
            && (normalized.Contains("無効", StringComparison.Ordinal)
                || normalized.Contains("invalid", StringComparison.OrdinalIgnoreCase)))
        {
            var componentName = componentMatch.Groups["name"].Value;
            insight = new ErrorInsight(
                LocalizeText("コンポーネント不一致", "Component Version Mismatch"),
                IsJapaneseUiLanguage() ? $"コンポーネント '{componentName}' は、このQAC環境では利用できません。" : $"Component '{componentName}' is not available in this QAC installation.",
                LocalizeText("qacli で確認できるコンポーネントバージョンに COM_* を合わせてください。", "Match COM_* values to the available component versions listed by qacli."));
            return true;
        }

        if (normalized.Contains("パースエラー", StringComparison.Ordinal)
            || normalized.Contains("parse error", StringComparison.OrdinalIgnoreCase))
        {
            var hint = normalized.Contains("-t (--type)", StringComparison.Ordinal)
                ? LocalizeText("qacli view -t の値を確認してください（ツールバージョンにより未対応の種別の可能性があります）。", "Review qacli view -t value (tool version may not support the given type).")
                : normalized.Contains("-m (--medium)", StringComparison.Ordinal)
                    ? LocalizeText("qacli view -m の値を確認してください（ツールバージョンで許容メディアが異なる可能性があります）。", "Review qacli view -m value (allowed media differ by tool version).")
                    : LocalizeText("インストール済みの qacli バージョンに対してコマンドオプションを見直してください。", "Review command options for the installed qacli version.");
            insight = new ErrorInsight(
                LocalizeText("CLIパースエラー", "CLI Parse Error"),
                LocalizeText("コマンド引数が現在の qacli では受け付けられません。", "Command arguments are not accepted by the current qacli."),
                hint);
            return true;
        }

        if ((normalized.Contains("見つけることができません", StringComparison.Ordinal)
                || normalized.Contains("見つかりません", StringComparison.Ordinal)
                || normalized.Contains("not found", StringComparison.OrdinalIgnoreCase))
            && (normalized.Contains(".h", StringComparison.OrdinalIgnoreCase)
                || normalized.Contains(".hpp", StringComparison.OrdinalIgnoreCase)))
        {
            insight = new ErrorInsight(
                LocalizeText("インクルードパスエラー", "Include Path Error"),
                LocalizeText("解析時に必要なヘッダーファイルを解決できませんでした。", "Required header file could not be resolved during analysis."),
                LocalizeText("プロジェクトの include path 設定と SOURCE_ROOT 解決結果を確認してください。", "Check project include path settings and SOURCE_ROOT resolution."));
            return true;
        }

        if (normalized.Contains("ユーザトークンを作成できません", StringComparison.Ordinal)
            || normalized.Contains("authentication", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("unauthorized", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("forbidden", StringComparison.OrdinalIgnoreCase))
        {
            insight = new ErrorInsight(
                LocalizeText("認証エラー", "Authentication Error"),
                LocalizeText("QAV/Validate の認証またはトークン設定に失敗しました。", "Authentication/token setup failed for QAV/Validate."),
                LocalizeText("サーバーURL・ユーザー・パスワード・権限スコープを確認してください。", "Check server URL, user, password, and permission scope."));
            return true;
        }

        if (normalized.Contains("[ERR]", StringComparison.OrdinalIgnoreCase)
            || IsFailureOutputLine(normalized)
            || IsErrorHintOutputLine(normalized))
        {
            var hint = InferConfigurationHint(normalized);
            if (string.IsNullOrWhiteSpace(hint))
            {
                hint = LocalizeText("参照された qacli ログを開き、このエラー直前のコマンドを確認してください。", "Open the referenced qacli log and check the command just before this error.");
            }

            insight = new ErrorInsight(
                LocalizeText("実行時エラー", "Runtime Error"),
                LocalizeText("テスト出力で実行時失敗の兆候を検出しました。", "A runtime failure indicator was detected in the test output."),
                hint);
            return true;
        }

        return false;
    }

    private void SetError(string message)
    {
        LastError = message;
        StatusText = message;
        ErrorLineCount += 1;
        AppendLog($"[ERR] {message}", true, false);
        _ = AppendRunLogLineAsync($"[ERR] {message}");
        RecordErrorFinding(message);
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

    private sealed record PrecheckProcessResult(
        int ExitCode,
        string StdOut,
        string StdErr,
        bool TimedOut);

    private sealed record ErrorInsight(
        string Category,
        string Explanation,
        string Hint);
}

public partial class ParsedErrorItemViewModel : ObservableObject
{
    public ParsedErrorItemViewModel(int index, string category, string errorText, string explanation, string hint)
    {
        Index = index;
        Category = category;
        ErrorText = errorText;
        Explanation = explanation;
        Hint = hint;
        FirstSeenTime = DateTime.Now.ToString("HH:mm:ss");
        LastSeenTime = FirstSeenTime;
        Occurrences = 1;
    }

    [ObservableProperty]
    private int index;

    [ObservableProperty]
    private string category;

    [ObservableProperty]
    private string errorText;

    [ObservableProperty]
    private string explanation;

    [ObservableProperty]
    private string hint;

    [ObservableProperty]
    private string firstSeenTime;

    [ObservableProperty]
    private string lastSeenTime;

    [ObservableProperty]
    private int occurrences;

    public void MarkOccurrence()
    {
        Occurrences += 1;
        LastSeenTime = DateTime.Now.ToString("HH:mm:ss");
    }
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
            ["Precheck"] = "Precheck",
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
            ["SectionErrorAnalysis"] = "Error Analysis",
            ["ExtractedErrorCount"] = "Extracted Errors",
            ["ErrorCategory"] = "Category",
            ["ErrorExplanation"] = "Explanation",
            ["ErrorHint"] = "Hint",
            ["ErrorOccurrences"] = "Occurrences",
            ["ErrorFirstSeen"] = "First Seen",
            ["ErrorLastSeen"] = "Last Seen",
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

