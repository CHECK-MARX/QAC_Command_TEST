using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using LinUCommandTestGui.ViewModels;

namespace LinUCommandTestGui.Views;

public partial class MainWindow : Window
{
    private bool _autoScrollScheduled;
    private MainWindowViewModel? _subscribedViewModel;
    private readonly DispatcherTimer _liveLogAutoScrollTimer = new() { Interval = TimeSpan.FromMilliseconds(500) };

    public MainWindow()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
        Closed += (_, _) => DetachLiveOutputSubscription();
        _liveLogAutoScrollTimer.Tick += (_, _) =>
        {
            if (DataContext is MainWindowViewModel { IsRunning: true, AutoScrollLiveOutput: true })
            {
                ScheduleLiveOutputScroll();
            }
        };
        _liveLogAutoScrollTimer.Start();
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

    private static string NormalizePickerPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        var trimmed = path.Trim().Trim('"');
        if (trimmed == "~")
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }

        if (trimmed.StartsWith("~/", StringComparison.Ordinal))
        {
            var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            return Path.Combine(home, trimmed[2..]);
        }

        return trimmed;
    }

    private void OnDataContextChanged(object? sender, EventArgs e)
    {
        AttachLiveOutputSubscription();
    }

    private void AttachLiveOutputSubscription()
    {
        DetachLiveOutputSubscription();

        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        _subscribedViewModel = vm;
        _subscribedViewModel.PropertyChanged += OnViewModelPropertyChanged;
        ScheduleLiveOutputScroll();
    }

    private void DetachLiveOutputSubscription()
    {
        if (_subscribedViewModel is not null)
        {
            _subscribedViewModel.PropertyChanged -= OnViewModelPropertyChanged;
            _subscribedViewModel = null;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        if (string.Equals(e.PropertyName, nameof(MainWindowViewModel.LiveOutputText), StringComparison.Ordinal))
        {
            ScheduleLiveOutputScroll();
            return;
        }

        if (string.Equals(e.PropertyName, nameof(MainWindowViewModel.AutoScrollLiveOutput), StringComparison.Ordinal)
            && vm.AutoScrollLiveOutput)
        {
            ScheduleLiveOutputScroll();
        }
    }

    private void ScheduleLiveOutputScroll()
    {
        if (DataContext is not MainWindowViewModel vm || !vm.AutoScrollLiveOutput)
        {
            return;
        }

        if (_autoScrollScheduled)
        {
            return;
        }

        _autoScrollScheduled = true;
        Dispatcher.UIThread.Post(() =>
        {
            _autoScrollScheduled = false;
            ScrollLiveOutputToBottom();
            Dispatcher.UIThread.Post(ScrollLiveOutputToBottom, DispatcherPriority.Render);
        }, DispatcherPriority.Background);
    }

    private void ScrollLiveOutputToBottom()
    {
        if (DataContext is not MainWindowViewModel vm || !vm.AutoScrollLiveOutput)
        {
            return;
        }

        if (string.IsNullOrEmpty(vm.LiveOutputText))
        {
            return;
        }

        LiveOutputTextBox.CaretIndex = vm.LiveOutputText.Length;
        var viewer = LiveOutputTextBox.GetVisualDescendants().OfType<ScrollViewer>().FirstOrDefault();
        if (viewer is null)
        {
            return;
        }

        var maxOffsetY = Math.Max(0, viewer.Extent.Height - viewer.Viewport.Height);
        viewer.Offset = new Vector(viewer.Offset.X, maxOffsetY);
    }

    private async void OnCopySelectedLiveOutputClick(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm)
        {
            return;
        }

        var selectedText = LiveOutputTextBox.SelectedText;
        if (string.IsNullOrEmpty(selectedText))
        {
            return;
        }

        try
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel?.Clipboard is null)
            {
                vm.ReportUnhandledException("Clipboard is not available.");
                return;
            }

            await topLevel.Clipboard.SetTextAsync(selectedText);
        }
        catch (Exception ex)
        {
            vm.ReportUnhandledException(ex.Message);
        }
    }
}
