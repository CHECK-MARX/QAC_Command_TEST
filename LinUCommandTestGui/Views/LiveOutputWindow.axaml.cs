using System;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using LinUCommandTestGui.ViewModels;

namespace LinUCommandTestGui.Views;

public partial class LiveOutputWindow : Window
{
    private bool _scrollScheduled;
    private MainWindowViewModel? _viewModel;

    public LiveOutputWindow()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => Subscribe();
        Closed += (_, _) => Unsubscribe();
    }

    private void Subscribe()
    {
        Unsubscribe();
        _viewModel = DataContext as MainWindowViewModel;
        if (_viewModel is not null)
        {
            _viewModel.PropertyChanged += OnViewModelPropertyChanged;
            ScheduleScroll();
        }
    }

    private void Unsubscribe()
    {
        if (_viewModel is not null)
        {
            _viewModel.PropertyChanged -= OnViewModelPropertyChanged;
            _viewModel = null;
        }
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainWindowViewModel.LiveOutputText)
            || (e.PropertyName == nameof(MainWindowViewModel.AutoScrollLiveOutput)
                && _viewModel?.AutoScrollLiveOutput == true))
        {
            ScheduleScroll();
        }
    }

    private void ScheduleScroll()
    {
        if (_viewModel?.AutoScrollLiveOutput != true || _scrollScheduled)
        {
            return;
        }

        _scrollScheduled = true;
        Dispatcher.UIThread.Post(() =>
        {
            _scrollScheduled = false;
            if (_viewModel?.AutoScrollLiveOutput != true || string.IsNullOrEmpty(_viewModel.LiveOutputText))
            {
                return;
            }

            LiveOutputTextBox.CaretIndex = _viewModel.LiveOutputText.Length;
            var viewer = LiveOutputTextBox.GetVisualDescendants().OfType<ScrollViewer>().FirstOrDefault();
            if (viewer is not null)
            {
                viewer.Offset = new Vector(viewer.Offset.X, Math.Max(0, viewer.Extent.Height - viewer.Viewport.Height));
            }
        }, DispatcherPriority.Render);
    }

    private async void OnCopySelectedLiveOutputClick(object? sender, RoutedEventArgs e)
    {
        if (_viewModel is null || string.IsNullOrEmpty(LiveOutputTextBox.SelectedText))
        {
            return;
        }

        try
        {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard is not null)
            {
                await clipboard.SetTextAsync(LiveOutputTextBox.SelectedText);
            }
        }
        catch (Exception ex)
        {
            _viewModel.ReportUnhandledException(ex.Message);
        }
    }
}
