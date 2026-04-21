namespace LinUCommandTestGui.ViewModels;

public sealed class LiveLogEntry
{
    public LiveLogEntry(string text, bool isError, bool isCommand)
    {
        Text = text;
        IsError = isError;
        IsCommand = isCommand;
    }

    public string Text { get; }

    public bool IsError { get; }

    public bool IsCommand { get; }
}
