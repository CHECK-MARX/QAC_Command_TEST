using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Interactivity;

namespace LinUCommandTestGui.Views;

public partial class ErrorAnalysisWindow : Window
{
    public ErrorAnalysisWindow()
    {
        InitializeComponent();
    }

    private void OnShowMainWindowClick(object? sender, RoutedEventArgs e)
    {
        if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime
            { MainWindow: MainWindow mainWindow })
        {
            mainWindow.RestoreFromDetachedWindow();
        }
    }
}
