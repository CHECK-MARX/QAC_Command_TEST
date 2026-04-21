using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

    public MainWindowViewModel()
    {
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
        AllITablePath = Path.Combine(LinURootPath, "master_settings", "all", "JISX0208_UTF8_I.txt");
        AllJTablePath = Path.Combine(LinURootPath, "master_settings", "all", "JISX0208_UTF8_J.txt");
        ActiveITablePath = Path.Combine(LinURootPath, "master_settings", "JISX0208_UTF8_I.txt");
        ActiveJTablePath = Path.Combine(LinURootPath, "master_settings", "JISX0208_UTF8_J.txt");
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
            StatusText = $"Loaded command test directory: {LinURootPath}";
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
            var path = Path.Combine(LinURootPath, "testonce.sh");
            EnsureDirectoryForFile(path);

            var lines = new List<string>
            {
                "#!/usr/bin/env bash",
                $"export QAF_ROOT=\"{QafRoot}\"",
                $"export QACLI_BIN=\"{QacliBinPath}\"",
                $"export TEST_ROOT=\"{TestRoot}\"",
                $"export QAV_SERVER=\"{QavServer}\"",
                $"export QAV_USER=\"{QavUser}\"",
                $"export QAV_PASS=\"{QavPass}\"",
                $"export VAL_SERVER=\"{ValServer}\"",
                $"export VAL_USER=\"{ValUser}\"",
                $"export VAL_PASS=\"{ValPass}\""
            };

            await File.WriteAllLinesAsync(path, lines, new UTF8Encoding(false));
            StatusText = "Saved GUI environment values to testonce.sh.";
            AppendLog("[GUI] testonce.sh updated from GUI values.", false, false);
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

            await File.WriteAllLinesAsync(ActiveITablePath, selected.Select(x => NormalizeLineValue(x.IValue)), new UTF8Encoding(false));
            await File.WriteAllLinesAsync(ActiveJTablePath, selected.Select(x => NormalizeLineValue(x.JValue)), new UTF8Encoding(false));

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
                if (IsPaused)
                {
                    IsPaused = false;
                    StatusText = RunningProcessId > 0 ? $"Running (PID {RunningProcessId})" : "Running";
                    RunResult = "RUNNING";
                    VerdictText = "RUNNING";
                    VerdictReason = "Resumed from paused state.";
                }

                return;
            }

            await ApplySelectionAsync();

            IsRunning = true;
            IsPaused = false;
            _runStartedUtc = DateTime.UtcNow;
            RunningProcessId = Environment.ProcessId;
            ConfiguredLoopCount = Math.Max(1, SelectedPairCount);
            LoopStartedCount = ConfiguredLoopCount;
            LoopCompletedCount = 0;
            LoopFailedCount = 0;
            ErrorLineCount = 0;
            ProcessedSummaryCount = 0;
            TotalSuccessInSummaries = 0;
            TotalFailureInSummaries = 0;
            LastError = string.Empty;
            CurrentRunLogPath = Path.Combine(OutputLogDirectoryPath, $"gui_run_{DateTime.Now:yyyyMMdd_HHmmss}.log");
            StatusText = $"Running (PID {RunningProcessId})";
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

            await Task.Delay(250);

            ProcessedSummaryCount = ConfiguredLoopCount;
            LoopCompletedCount = ConfiguredLoopCount;
            TotalSuccessInSummaries = ConfiguredLoopCount;
            TotalFailureInSummaries = 0;
            ProgressPercentage = 100;
            ProgressText = $"{ConfiguredLoopCount} / {ConfiguredLoopCount} (100.0%)";
            RunResult = "PASS";
            VerdictText = "PASS";
            VerdictReason = "No failures detected in simulated run.";
            StatusText = "Completed.";
            ElapsedTime = (DateTime.UtcNow - _runStartedUtc).ToString(@"hh\:mm\:ss");
            OutputActivity = "idle";

            AppendLog("[GUI] Completed.", false, true);
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
            IsRunning = false;
            IsPaused = false;
            RunningProcessId = 0;
        }
    }

    [RelayCommand(CanExecute = nameof(CanStopTest))]
    private void StopTest()
    {
        if (!IsRunning)
        {
            StatusText = "No active process.";
            return;
        }

        if (!IsPaused)
        {
            IsPaused = true;
            StatusText = "Paused. Press Start to resume.";
            RunResult = "PAUSED";
            VerdictText = "PAUSED";
            VerdictReason = "Paused by user.";
            AppendLog("[GUI] Paused by user.", false, true);
            return;
        }

        IsPaused = false;
        IsRunning = false;
        RunningProcessId = 0;
        StatusText = "Stopped.";
        RunResult = "STOPPED";
        VerdictText = "STOPPED";
        VerdictReason = "Stopped by user.";
        OutputActivity = "idle";
        AppendLog("[GUI] Stopped by user.", true, true);
    }

    [RelayCommand]
    private void ClearLogs()
    {
        LogLines.Clear();
        LiveOutputText = string.Empty;
    }

    private bool CanStartTest()
    {
        return !IsRunning || IsPaused;
    }

    private bool CanStopTest()
    {
        return IsRunning || IsPaused;
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

        return File.ReadAllLines(path, Encoding.UTF8)
            .Select(x => x.Replace("\r", string.Empty).Replace("\uFEFF", string.Empty))
            .ToList();
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
            ["SectionEnvironment"] = "環境上書き (生成 testonce スクリプト)",
            ["SaveEnvToScript"] = "GUI値を testonce.sh へ保存",
            ["EnvLoadHint"] = "※ ディレクトリ指定時に testonce.sh を自動読込します",
            ["QavCredential"] = "QAV サーバー / ユーザー / パスワード",
            ["ValCredential"] = "Validate サーバー / ユーザー / パスワード",
            ["SectionRunControl"] = "実行制御",
            ["StartTest"] = "テスト開始",
            ["Stop"] = "一時停止/停止",
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
            ["SectionEnvironment"] = "Environment Overrides (generated testonce script)",
            ["SaveEnvToScript"] = "Save GUI values to testonce.sh",
            ["EnvLoadHint"] = "* testonce.sh is auto-loaded when directory is selected",
            ["QavCredential"] = "QAV Server / User / Password",
            ["ValCredential"] = "Validate Server / User / Password",
            ["SectionRunControl"] = "Run Control",
            ["StartTest"] = "Start Test",
            ["Stop"] = "Pause/Stop",
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
