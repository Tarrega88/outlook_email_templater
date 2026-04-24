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
    private ResolveState _state = ResolveState.Empty;
    private string _query = "";
    public string Name { get => _name; set { if (Set(ref _name, value)) { Raise(nameof(Display)); Raise(nameof(PillName)); } } }
    public string Email { get => _email; set { if (Set(ref _email, value)) { Raise(nameof(Display)); Raise(nameof(PillName)); } } }
    public bool IsCc { get => _isCc; set => Set(ref _isCc, value); }

    /// <summary>UI-only. Not serialized. Drives the row's status icon and pill styling.</summary>
    [System.Text.Json.Serialization.JsonIgnore]
    public ResolveState State { get => _state; set { if (Set(ref _state, value)) { Raise(nameof(IsResolved)); Raise(nameof(IsResolving)); Raise(nameof(IsUnresolved)); Raise(nameof(IsUnverified)); Raise(nameof(IsFromOutlook)); Raise(nameof(IsFromLocal)); } } }

    [System.Text.Json.Serialization.JsonIgnore] public bool IsResolved   => _state == ResolveState.ResolvedOutlook || _state == ResolveState.ResolvedLocal || _state == ResolveState.Unverified;
    [System.Text.Json.Serialization.JsonIgnore] public bool IsResolving  => _state == ResolveState.Resolving;
    [System.Text.Json.Serialization.JsonIgnore] public bool IsUnresolved => _state == ResolveState.Unresolved;
    [System.Text.Json.Serialization.JsonIgnore] public bool IsUnverified => _state == ResolveState.Unverified;
    [System.Text.Json.Serialization.JsonIgnore] public bool IsFromOutlook => _state == ResolveState.ResolvedOutlook;
    [System.Text.Json.Serialization.JsonIgnore] public bool IsFromLocal   => _state == ResolveState.ResolvedLocal;

    public string Display => string.IsNullOrWhiteSpace(Name) ? Email : $"{Name} <{Email}>";

    /// <summary>What to show as the big label of the pill. Prefers Name, else Email.</summary>
    public string PillName => string.IsNullOrWhiteSpace(Name) ? Email : Name;

    /// <summary>
    /// Transient text the user is editing in the recipient field. Not persisted.
    /// The pill swaps back to the TextBox whenever <see cref="State"/> leaves Resolved,
    /// at which point the TextBox is populated from this property.
    /// </summary>
    [System.Text.Json.Serialization.JsonIgnore]
    public string Query { get => _query; set => Set(ref _query, value); }
}

public enum ResolveState
{
    /// <summary>No text entered.</summary>
    Empty,
    /// <summary>Text is being resolved via Outlook right now.</summary>
    Resolving,
    /// <summary>Matched a contact that originally came from Outlook.</summary>
    ResolvedOutlook,
    /// <summary>Matched an entry in the local cache only (never seen in Outlook).</summary>
    ResolvedLocal,
    /// <summary>Valid email syntax but not in Outlook or the cache — send at your own risk.</summary>
    Unverified,
    /// <summary>Not a known contact and not a valid literal email.</summary>
    Unresolved
}

public class Contact : NotifyBase
{
    private string _name = "";
    private string _email = "";
    private DateTime _lastUsedUtc;
    private bool _foundInOutlook;
    public string Name { get => _name; set => Set(ref _name, value); }
    public string Email { get => _email; set => Set(ref _email, value); }

    /// <summary>When this contact was last picked/resolved. Used to rank ambiguous matches.</summary>
    public DateTime LastUsedUtc { get => _lastUsedUtc; set => Set(ref _lastUsedUtc, value); }

    /// <summary>
    /// True if this contact was ever confirmed by Outlook's address book resolver.
    /// Once true, stays true — Outlook-confirmed identities are more trustworthy than
    /// cache-only ones.
    /// </summary>
    public bool FoundInOutlook { get => _foundInOutlook; set => Set(ref _foundInOutlook, value); }

    /// <summary>
    /// Strings that have previously resolved to this contact (full name, first name,
    /// local-part, nicknames). All matching is case-insensitive and deduped on add.
    /// </summary>
    public List<string> Aliases { get; set; } = new();
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

public class EmailTemplate : NotifyBase
{
    private string _name = "New email";
    private string _subject = "";
    private string _body = "";

    public string Name { get => _name; set => Set(ref _name, value); }
    public string Subject { get => _subject; set => Set(ref _subject, value); }
    public string Body { get => _body; set => Set(ref _body, value); }

    public ObservableCollection<Recipient> Recipients
    {
        get => _recipients;
        set
        {
            if (_recipients != null) _recipients.CollectionChanged -= OnRecipientsChanged;
            _recipients = value ?? new ObservableCollection<Recipient>();
            _recipients.CollectionChanged += OnRecipientsChanged;
            foreach (var r in _recipients) r.PropertyChanged += OnRecipientChanged;
            Raise(nameof(Recipients));
            Raise(nameof(FilledRecipientCount));
        }
    }
    private ObservableCollection<Recipient> _recipients = new();
    public ObservableCollection<AttachmentItem> Attachments { get; set; } = new();

    public EmailTemplate()
    {
        _recipients.CollectionChanged += OnRecipientsChanged;
    }

    /// <summary>Number of recipient rows that have a non-blank email address.</summary>
    [System.Text.Json.Serialization.JsonIgnore]
    public int FilledRecipientCount =>
        Recipients.Count(r => !string.IsNullOrWhiteSpace(r.Email));

    private void OnRecipientsChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
    {
        if (e.OldItems != null)
            foreach (Recipient r in e.OldItems) r.PropertyChanged -= OnRecipientChanged;
        if (e.NewItems != null)
            foreach (Recipient r in e.NewItems) r.PropertyChanged += OnRecipientChanged;
        Raise(nameof(FilledRecipientCount));
    }

    private void OnRecipientChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(Recipient.Email))
            Raise(nameof(FilledRecipientCount));
    }
}

public class AppConfig : NotifyBase
{
    private string _senderName = "";
    private SendMode _sendMode = SendMode.OpenDraft;
    private bool _htmlBody = false;

    public string SenderName { get => _senderName; set => Set(ref _senderName, value); }
    public SendMode SendMode { get => _sendMode; set => Set(ref _sendMode, value); }
    public bool HtmlBody { get => _htmlBody; set => Set(ref _htmlBody, value); }

    public ObservableCollection<EmailTemplate> Emails { get; set; } = new();

    /// <summary>
    /// In-memory contact book, shared across all configs via <c>ContactsService</c>.
    /// Not serialized into per-config JSON files — contacts live in their own
    /// file so opening a different config doesn't wipe or duplicate them.
    /// </summary>
    [System.Text.Json.Serialization.JsonIgnore]
    public ObservableCollection<Contact> Contacts { get; set; } = new();

    /// <summary>
    /// Legacy slot: older config files embedded contacts under "Contacts". We
    /// deserialize those into here and migrate them into the shared store on load.
    /// Set to <c>null</c> after migration so it's omitted from re-saved files.
    /// </summary>
    [System.Text.Json.Serialization.JsonPropertyName("Contacts")]
    public List<Contact>? LegacyContacts { get; set; }
}
