namespace LinUCommandTestGui;

// A delayed automatic Enter and a manual click may refer to the same prompt.
// Only one may consume it; an old request must never consume a newer prompt.
public sealed class PromptConfirmationGate
{
    public long Generation { get; private set; }
    private bool pending;

    public void Observe()
    {
        Generation++;
        pending = true;
    }

    public void Clear()
    {
        Generation++;
        pending = false;
    }

    public bool TryConsume(long generation)
    {
        if (!pending || generation != Generation) return false;
        pending = false;
        return true;
    }
}
