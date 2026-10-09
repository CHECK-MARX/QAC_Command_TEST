using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using LinUCommandTestGui.ViewModels;

namespace LinUCommandTestGui.Views;

public partial class MainWindow : Window
{
    private LiveOutputWindow? _liveOutputWindow;
    private ErrorAnalysisWindow? _errorAnalysisWindow;

    public MainWindow()
    {
        InitializeComponent();
    }

    private void OnOpenLiveOutputWindowClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        if (_liveOutputWindow is null || !_liveOutputWindow.IsVisible)
        {
            _liveOutputWindow = new LiveOutputWindow { DataContext = vm };
            _liveOutputWindow.Closed += (_, _) =>
            {
                _liveOutputWindow = null;
                UpdateDetachedWindowState();
            };
            if (OperatingSystem.IsLinux())
            {
                _liveOutputWindow.Show();
            }
            else
            {
                _liveOutputWindow.Show(this);
            }
            UpdateDetachedWindowState();
        }
        else
        {
            if (OperatingSystem.IsLinux() && _liveOutputWindow.WindowState == WindowState.Minimized)
            {
                _liveOutputWindow.WindowState = WindowState.Normal;
            }

            _liveOutputWindow.Activate();
        }
    }

    private void OnOpenErrorAnalysisWindowClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        if (_errorAnalysisWindow is null || !_errorAnalysisWindow.IsVisible)
        {
            _errorAnalysisWindow = new ErrorAnalysisWindow { DataContext = vm };
            _errorAnalysisWindow.Closed += (_, _) =>
            {
                _errorAnalysisWindow = null;
                UpdateDetachedWindowState();
            };
            if (OperatingSystem.IsLinux())
            {
                _errorAnalysisWindow.Show();
            }
            else
            {
                _errorAnalysisWindow.Show(this);
            }
            UpdateDetachedWindowState();
        }
        else
        {
            if (OperatingSystem.IsLinux() && _errorAnalysisWindow.WindowState == WindowState.Minimized)
            {
                _errorAnalysisWindow.WindowState = WindowState.Normal;
            }

            _errorAnalysisWindow.Activate();
        }
    }

    private void OnHideMainWindowClick(object? sender, RoutedEventArgs e)
    {
        if (OperatingSystem.IsLinux()
            && (_liveOutputWindow is not null || _errorAnalysisWindow is not null))
        {
            Hide();
        }
    }

    private void UpdateDetachedWindowState()
    {
        if (!OperatingSystem.IsLinux())
        {
            return;
        }

        var hasDetachedWindow = _liveOutputWindow is not null || _errorAnalysisWindow is not null;
        HideMainWindowButton.IsEnabled = hasDetachedWindow;
        if (!hasDetachedWindow && !IsVisible)
        {
            RestoreFromDetachedWindow();
        }
    }

    public void RestoreFromDetachedWindow()
    {
        if (!OperatingSystem.IsLinux())
        {
            return;
        }

        if (!IsVisible)
        {
            Show();
        }

        if (WindowState == WindowState.Minimized)
        {
            WindowState = WindowState.Normal;
        }

        Activate();
    }

    private async void OnBrowseCommandTestDirectoryClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel?.StorageProvider is null)
            {
                vm.ReportUnhandledException("Storage provider is not available.");
                return;
            }

            var options = new FolderPickerOpenOptions
            {
                Title = "Select command test directory",
                AllowMultiple = false
            };

            var normalized = NormalizePickerPath(vm.CommandTestDirectoryPath);
            if (Directory.Exists(normalized))
            {
                var suggested = await topLevel.StorageProvider.TryGetFolderFromPathAsync(normalized);
                if (suggested is not null)
                {
                    options.SuggestedStartLocation = suggested;
                }
            }

            var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(options);
            var selected = folders.FirstOrDefault();
            if (selected is null)
            {
                return;
            }

            var localPath = selected.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(localPath) && selected.Path.IsAbsoluteUri)
            {
                localPath = Uri.UnescapeDataString(selected.Path.LocalPath);
            }

            if (string.IsNullOrWhiteSpace(localPath))
            {
                vm.ReportUnhandledException("Failed to resolve selected folder path.");
                return;
            }

            vm.CommandTestDirectoryPath = localPath;
            await vm.LoadFromCommandTestDirectoryCommand.ExecuteAsync(null);
        }
        catch (Exception ex)
        {
            vm.ReportUnhandledException(ex.Message);
        }
    }

    private async void OnBrowseEnvironmentDirectoryClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        if (sender is not Button { Tag: string targetField })
        {
            return;
        }

        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel?.StorageProvider is null)
            {
                vm.ReportUnhandledException("Storage provider is not available.");
                return;
            }

            var currentPath = GetEnvironmentDirectoryValue(vm, targetField);
            var options = new FolderPickerOpenOptions
            {
                Title = "Select directory",
                AllowMultiple = false
            };

            var normalized = NormalizePickerPath(currentPath, vm);
            if (Directory.Exists(normalized))
            {
                var suggested = await topLevel.StorageProvider.TryGetFolderFromPathAsync(normalized);
                if (suggested is not null)
                {
                    options.SuggestedStartLocation = suggested;
                }
            }

            var folders = await topLevel.StorageProvider.OpenFolderPickerAsync(options);
            var selected = folders.FirstOrDefault();
            if (selected is null)
            {
                return;
            }

            var localPath = selected.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(localPath) && selected.Path.IsAbsoluteUri)
            {
                localPath = Uri.UnescapeDataString(selected.Path.LocalPath);
            }

            if (string.IsNullOrWhiteSpace(localPath))
            {
                vm.ReportUnhandledException("Failed to resolve selected folder path.");
                return;
            }

            SetEnvironmentDirectoryValue(vm, targetField, localPath);
        }
        catch (Exception ex)
        {
            vm.ReportUnhandledException(ex.Message);
        }
    }

    private static string GetEnvironmentDirectoryValue(MainWindowViewModel vm, string targetField)
    {
        return targetField switch
        {
            nameof(MainWindowViewModel.QafRoot) => vm.QafRoot,
            nameof(MainWindowViewModel.QacliBinPath) => vm.QacliBinPath,
            nameof(MainWindowViewModel.TestRoot) => vm.TestRoot,
            _ => string.Empty
        };
    }

    private static void SetEnvironmentDirectoryValue(MainWindowViewModel vm, string targetField, string value)
    {
        switch (targetField)
        {
            case nameof(MainWindowViewModel.QafRoot):
                vm.QafRoot = value;
                break;
            case nameof(MainWindowViewModel.QacliBinPath):
                vm.QacliBinPath = value;
                break;
            case nameof(MainWindowViewModel.TestRoot):
                vm.TestRoot = value;
                break;
        }
    }

    private static string NormalizePickerPath(string path, MainWindowViewModel? vm = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        var trimmed = path.Trim().Trim('"').Replace("%QAF_ROOT%", vm?.QafRoot ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        trimmed = Environment.ExpandEnvironmentVariables(trimmed);
        if (trimmed == "~")
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }

        if (trimmed.StartsWith("~/", StringComparison.Ordinal))
        {
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(home, trimmed[2..]);
        }

        if (!Path.IsPathRooted(trimmed) && vm is not null && !string.IsNullOrWhiteSpace(vm.LinURootPath))
        {
            trimmed = Path.Combine(vm.LinURootPath, trimmed);
        }

        return trimmed;
    }

}
