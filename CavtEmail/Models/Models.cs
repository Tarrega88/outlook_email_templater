using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CavtEmail.Models;

public abstract class NotifyBase : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected bool Set<T>(ref T field, T value, [CallerMemberName] string? prop = null)
    {
        if (Equals(field, value)) return false;
        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        return true;
    }

    protected void Raise([CallerMemberName] string? prop = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
}

public class Recipient : NotifyBase
{
    private string _name = "";
    private string _email = "";
    private bool _isCc;
    public string Name { get => _name; set => Set(ref _name, value); }
    public string Email { get => _email; set => Set(ref _email, value); }
    public bool IsCc { get => _isCc; set => Set(ref _isCc, value); }

    public string Display => string.IsNullOrWhiteSpace(Name) ? Email : $"{Name} <{Email}>";
}

public class Contact : NotifyBase
{
    private string _name = "";
    private string _email = "";
    public string Name { get => _name; set => Set(ref _name, value); }
    public string Email { get => _email; set => Set(ref _email, value); }
}

public class AttachmentItem : NotifyBase
{
    private string _path = "";
    public string Path
    {
        get => _path;
        set
        {
            if (Set(ref _path, value))
            {
                Raise(nameof(FileName));
                Raise(nameof(Exists));
                Raise(nameof(IsFolder));
                Raise(nameof(DisplayKind));
                Raise(nameof(ResolvedFileCount));
            }
        }
    }

    public string FileName => System.IO.Path.GetFileName(Path.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar));

    /// <summary>True when <see cref="Path"/> points to a directory that currently exists.</summary>
    public bool IsFolder => !string.IsNullOrWhiteSpace(Path) && System.IO.Directory.Exists(Path);

    /// <summary>True when the path exists as either a file or a folder.</summary>
    public bool Exists => System.IO.File.Exists(Path) || System.IO.Directory.Exists(Path);

    /// <summary>Short label used in the UI to distinguish folders from files.</summary>
    public string DisplayKind => IsFolder ? "📁 Folder" : "";

    /// <summary>For folder refs, the live top-level file count (best-effort). Empty string for files.</summary>
    public string ResolvedFileCount
    {
        get
        {
            if (!IsFolder) return "";
            try
            {
                var n = System.IO.Directory.EnumerateFiles(Path).Count();
                return n == 1 ? "1 file" : $"{n} files";
            }
            catch
            {
                return "inaccessible";
            }
        }
    }

    /// <summary>
    /// Resolves this attachment to the actual file paths to send.
    /// For a file, yields the single path (if it exists). For a folder, yields
    /// the top-level files inside it (non-recursive) at the moment of the call.
    /// </summary>
    public IEnumerable<string> ResolveFiles()
    {
        if (string.IsNullOrWhiteSpace(Path)) yield break;

        if (System.IO.Directory.Exists(Path))
        {
            IEnumerable<string> files;
            try { files = System.IO.Directory.EnumerateFiles(Path); }
            catch { yield break; }
            foreach (var f in files) yield return f;
        }
        else if (System.IO.File.Exists(Path))
        {
            yield return Path;
        }
    }
}

public enum SendMode
{
    OpenDraft = 0,
    SendAutomatically = 1
}

public class EmailGroup : NotifyBase
{
    private string _name = "New Group";
    private string _subject = "";
    private string _body = "";

    public string Name { get => _name; set => Set(ref _name, value); }
    public string Subject { get => _subject; set => Set(ref _subject, value); }
    public string Body { get => _body; set => Set(ref _body, value); }

    public ObservableCollection<Recipient> Recipients { get; set; } = new();
    public ObservableCollection<AttachmentItem> Attachments { get; set; } = new();
}

public class AppConfig : NotifyBase
{
    private string _senderName = "";
    private SendMode _sendMode = SendMode.OpenDraft;
    private bool _htmlBody = false;

    public string SenderName { get => _senderName; set => Set(ref _senderName, value); }
    public SendMode SendMode { get => _sendMode; set => Set(ref _sendMode, value); }
    public bool HtmlBody { get => _htmlBody; set => Set(ref _htmlBody, value); }

    public ObservableCollection<EmailGroup> Groups { get; set; } = new();
    public ObservableCollection<Contact> Contacts { get; set; } = new();
}
