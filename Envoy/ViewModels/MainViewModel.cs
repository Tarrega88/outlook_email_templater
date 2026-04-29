using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using Envoy.Models;
using Envoy.Services;
using Microsoft.Win32;

namespace Envoy.ViewModels;

public class MainViewModel : NotifyBase
{
    private AppConfig _config = new();
    private EmailTemplate? _selectedEmail;
    private string _currentConfigPath;
    private string _statusMessage = "";
    private bool _isOutlookAvailable;
    private readonly DispatcherTimer _outlookPingTimer;

    // JSON snapshot of the config as last loaded/saved. The dirty-check timer
    // re-serializes the live config and compares against this string to flip
    // IsDirty without us having to wire change-tracking into every model.
    private string _savedSnapshot = "";
    private readonly DispatcherTimer _dirtyCheckTimer;
    private bool _isDirty;

    public MainViewModel()
    {
        _currentConfigPath = ConfigService.DefaultPath;
        _config = ConfigService.Load(_currentConfigPath);

        // Contacts live in a shared file independent of whatever config is open.
        _config.Contacts = ContactsService.Load();
        MigrateLegacyContacts(_config);

        if (_config.Emails.Count == 0)
        {
            _config.Emails.Add(new EmailTemplate { Name = "Email 1" });
        }
        StampRecipientStates(_config);
        foreach (var e in _config.Emails) EnsureTrailingBlankRecipient(e);
        _selectedEmail = _config.Emails.FirstOrDefault();

        // First-run: auto-fill the sender display name from Windows / AD.
        if (string.IsNullOrWhiteSpace(_config.SenderName))
        {
            _config.SenderName = CurrentUserInfo.TryGetDisplayName()
                                 ?? Environment.UserName
                                 ?? "";
        }

        Resolver = new ContactResolver(_config.Contacts, SaveContacts);

        AddEmailCommand          = new RelayCommand(AddEmail);
        RemoveEmailCommand       = new RelayCommand(RemoveEmail, () => SelectedEmail != null);
        DuplicateEmailCommand    = new RelayCommand(DuplicateEmail, () => SelectedEmail != null);

        AddRecipientCommand      = new RelayCommand(AddRecipient, () => SelectedEmail != null);
        RemoveRecipientCommand   = new RelayCommand(RemoveRecipient);

        AddAttachmentCommand        = new RelayCommand(AddAttachment, () => SelectedEmail != null);
        AddAttachmentFolderCommand  = new RelayCommand(AddAttachmentFolder, () => SelectedEmail != null);
        RemoveAttachmentCommand     = new RelayCommand(RemoveAttachment);

        SaveCommand              = new RelayCommand(Save);
        SaveAsCommand            = new RelayCommand(SaveAs);
        OpenCommand              = new RelayCommand(Open);

        SendCommand              = new RelayCommand(Send, () => SelectedEmail != null);
        SendAllCommand           = new RelayCommand(SendAll, () => _config.Emails.Count > 0);
        PreviewCommand           = new RelayCommand(Preview, () => SelectedEmail != null);

        ReVerifyContactsCommand  = new RelayCommand(ReVerifyContacts, () => _config.Contacts.Count > 0);
        ClearContactsCommand     = new RelayCommand(ClearContacts,    () => _config.Contacts.Count > 0);

        // Poll Outlook availability for the toolbar indicator. Cheap (process list),
        // no COM launch. 3s interval keeps the indicator responsive without churn.
        IsOutlookAvailable = ContactResolver.IsOutlookRunning();
        _outlookPingTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
        _outlookPingTimer.Tick += (_, _) => IsOutlookAvailable = ContactResolver.IsOutlookRunning();
        _outlookPingTimer.Start();

        // Take an initial clean snapshot, then poll for changes for the title-bar
        // dirty marker. 500ms is plenty responsive for typing without churn.
        SnapshotSaved();
        _dirtyCheckTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        _dirtyCheckTimer.Tick += (_, _) => RecomputeDirty();
        _dirtyCheckTimer.Start();
    }

    private void SnapshotSaved()
    {
        try { _savedSnapshot = ConfigService.Serialize(_config); }
        catch { _savedSnapshot = ""; }
        IsDirty = false;
    }

    private void RecomputeDirty()
    {
        try
        {
            var current = ConfigService.Serialize(_config);
            IsDirty = !string.Equals(current, _savedSnapshot, StringComparison.Ordinal);
        }
        catch { /* ignore — leave IsDirty as-is */ }
    }

    public ContactResolver Resolver { get; }

    public AppConfig Config { get => _config; set { Set(ref _config, value); Raise(nameof(Emails)); } }
    public ObservableCollection<EmailTemplate> Emails => _config.Emails;

    public EmailTemplate? SelectedEmail
    {
        get => _selectedEmail;
        set { if (Set(ref _selectedEmail, value)) Raise(nameof(HasSelection)); }
    }

    public bool HasSelection => _selectedEmail != null;

    public string StatusMessage { get => _statusMessage; set => Set(ref _statusMessage, value); }

    public string CurrentConfigPath
    {
        get => _currentConfigPath;
        set { if (Set(ref _currentConfigPath, value)) Raise(nameof(WindowTitle)); }
    }

    public bool IsDirty
    {
        get => _isDirty;
        private set { if (Set(ref _isDirty, value)) Raise(nameof(WindowTitle)); }
    }

    /// <summary>
    /// Window chrome title: <c>Envoy — &lt;name&gt;</c>, with a trailing <c>*</c>
    /// when there are unsaved changes. Falls back to <c>Untitled</c> for an
    /// unsaved/blank config path.
    /// </summary>
    public string WindowTitle
    {
        get
        {
            string name;
            try { name = Path.GetFileNameWithoutExtension(_currentConfigPath); }
            catch { name = ""; }
            if (string.IsNullOrWhiteSpace(name)) name = "Untitled";
            return $"Envoy \u2014 {name}{(_isDirty ? "*" : "")}";
        }
    }

    public Array SendModes { get; } = Enum.GetValues(typeof(SendMode));

    public RelayCommand AddEmailCommand { get; }
    public RelayCommand RemoveEmailCommand { get; }
    public RelayCommand DuplicateEmailCommand { get; }
    public RelayCommand AddRecipientCommand { get; }
    public RelayCommand RemoveRecipientCommand { get; }
    public RelayCommand AddAttachmentCommand { get; }
    public RelayCommand AddAttachmentFolderCommand { get; }
    public RelayCommand RemoveAttachmentCommand { get; }
    public RelayCommand SaveCommand { get; }
    public RelayCommand SaveAsCommand { get; }
    public RelayCommand OpenCommand { get; }
    public RelayCommand SendCommand { get; }
    public RelayCommand SendAllCommand { get; }
    public RelayCommand ReVerifyContactsCommand { get; }
    public RelayCommand ClearContactsCommand { get; }
    public RelayCommand PreviewCommand { get; }

    public bool IsOutlookAvailable
    {
        get => _isOutlookAvailable;
        private set => Set(ref _isOutlookAvailable, value);
    }

    public ObservableCollection<Contact> Contacts => _config.Contacts;

    /// <summary>Persist the shared contacts file. Safe to call freely.</summary>
    public void SaveContacts()
    {
        try { ContactsService.Save(_config.Contacts); }
        catch (Exception ex) { StatusMessage = "Contacts save failed: " + ex.Message; }
    }

    /// <summary>
    /// When a config is loaded, categorize each recipient against the current
    /// contact cache so the UI shows the right provenance icon (Outlook vs.
    /// local vs. unverified literal) without forcing a re-resolve.
    /// </summary>
    private void StampRecipientStates(AppConfig cfg)
    {
        foreach (var e in cfg.Emails)
        foreach (var r in e.Recipients)
        {
            if (string.IsNullOrWhiteSpace(r.Email))
            {
                r.State = ResolveState.Empty;
                r.Query = "";
                continue;
            }

            var hit = cfg.Contacts.FirstOrDefault(c =>
                string.Equals(c.Email, r.Email, StringComparison.OrdinalIgnoreCase));
            if (hit != null)
            {
                r.State = hit.FoundInOutlook
                    ? ResolveState.ResolvedOutlook
                    : ResolveState.ResolvedLocal;
            }
            else
            {
                // Address is only known to this config. Treat as local-but-unverified:
                // valid email -> Unverified, otherwise Unresolved.
                r.State = ContactResolver.LooksLikeEmail(r.Email)
                    ? ResolveState.Unverified
                    : ResolveState.Unresolved;
            }
            r.Query = string.IsNullOrWhiteSpace(r.Name) ? r.Email : r.Name;
        }
    }

    /// <summary>
    /// Ensure <paramref name="e"/> has a trailing blank recipient row so the user
    /// never has to click "+ Add" just to type the first (or next) address.
    /// </summary>
    public void EnsureTrailingBlankRecipient(EmailTemplate? e)
    {
        if (e == null) return;
        var last = e.Recipients.Count == 0 ? null : e.Recipients[^1];
        if (last == null ||
            !string.IsNullOrWhiteSpace(last.Name) ||
            !string.IsNullOrWhiteSpace(last.Email))
        {
            e.Recipients.Add(new Recipient { Name = "", Email = "" });
        }
    }

    /// <summary>
    /// If an opened config carried an embedded "Contacts" array (old format),
    /// fold those entries into the shared store and persist.
    /// </summary>
    private void MigrateLegacyContacts(AppConfig cfg)
    {
        if (cfg.LegacyContacts == null || cfg.LegacyContacts.Count == 0) return;
        var added = ContactsService.Merge(cfg.Contacts, cfg.LegacyContacts);
        cfg.LegacyContacts = null;
        if (added > 0) SaveContacts();
    }

    // --------------------------------------------------------------
    // Contacts menu actions
    // --------------------------------------------------------------

    private void ClearContacts()
    {
        var count = _config.Contacts.Count;
        if (count == 0) return;

        var result = MessageBox.Show(
            $"Permanently delete all {count} saved contact{(count == 1 ? "" : "s")}?\r\n\r\n" +
            "This cannot be undone. Contacts you've added manually will be lost; " +
            "Outlook address book entries will be re-discovered the next time you type them.",
            "Clear all contacts",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning,
            MessageBoxResult.No);

        if (result != MessageBoxResult.Yes) return;

        Resolver.ClearAll();
        SaveContacts();
        StatusMessage = $"Cleared {count} contact{(count == 1 ? "" : "s")}.";
    }

    private void ReVerifyContacts()
    {
        if (!ContactResolver.IsOutlookRunning())
        {
            MessageBox.Show(
                "Classic Outlook must be running to re-verify contacts against the address book.",
                "Outlook not running", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        StatusMessage = "Re-verifying contacts against Outlook…";
        var report = Resolver.ReVerifyAgainstOutlook();
        SaveContacts();

        StatusMessage =
            $"Re-verified {report.Checked} contact{(report.Checked == 1 ? "" : "s")}: " +
            $"{report.Promoted} newly verified, {report.Demoted} no longer in Outlook, {report.Unchanged} unchanged.";
    }

    private void AddEmail()
    {
        var e = new EmailTemplate { Name = $"Email {_config.Emails.Count + 1}" };
        _config.Emails.Add(e);
        EnsureTrailingBlankRecipient(e);
        SelectedEmail = e;
    }

    private void RemoveEmail()
    {
        if (SelectedEmail == null) return;
        if (MessageBox.Show($"Delete email '{SelectedEmail.Name}'?", "Confirm",
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        var idx = _config.Emails.IndexOf(SelectedEmail);
        _config.Emails.Remove(SelectedEmail);
        SelectedEmail = _config.Emails.Count == 0
            ? null
            : _config.Emails[Math.Min(idx, _config.Emails.Count - 1)];
    }

    private void DuplicateEmail()
    {
        if (SelectedEmail == null) return;
        var src = SelectedEmail;
        var copy = new EmailTemplate
        {
            Name = src.Name + " (copy)",
            Subject = src.Subject,
            Body = src.Body,
        };
        foreach (var r in src.Recipients) copy.Recipients.Add(new Recipient { Name = r.Name, Email = r.Email, IsCc = r.IsCc });
        foreach (var a in src.Attachments) copy.Attachments.Add(new AttachmentItem { Path = a.Path });
        _config.Emails.Add(copy);
        EnsureTrailingBlankRecipient(copy);
        SelectedEmail = copy;
    }

    private void AddRecipient()
    {
        SelectedEmail?.Recipients.Add(new Recipient { Name = "", Email = "" });
    }

    private void RemoveRecipient(object? param)
    {
        if (SelectedEmail != null && param is Recipient r)
        {
            SelectedEmail.Recipients.Remove(r);
            EnsureTrailingBlankRecipient(SelectedEmail);
        }
    }

    private void AddAttachment()
    {
        if (SelectedEmail == null) return;
        var dlg = new OpenFileDialog { Multiselect = true, Title = "Select files to attach" };
        if (dlg.ShowDialog() == true)
        {
            foreach (var f in dlg.FileNames)
                AddAttachmentPath(f);
        }
    }

    private void AddAttachmentFolder()
    {
        if (SelectedEmail == null) return;
        var dlg = new OpenFolderDialog
        {
            Title = "Select a folder — its top-level files will be attached at send time",
            Multiselect = true
        };
        if (dlg.ShowDialog() == true)
        {
            foreach (var f in dlg.FolderNames)
                AddAttachmentPath(f);
        }
    }

    public void AddAttachmentPath(string path)
    {
        if (SelectedEmail == null || string.IsNullOrWhiteSpace(path)) return;
        if (SelectedEmail.Attachments.Any(a => string.Equals(a.Path, path, StringComparison.OrdinalIgnoreCase))) return;
        SelectedEmail.Attachments.Add(new AttachmentItem { Path = path });
    }

    private void RemoveAttachment(object? param)
    {
        if (SelectedEmail != null && param is AttachmentItem a)
            SelectedEmail.Attachments.Remove(a);
    }

    private void Save()
    {
        try
        {
            ConfigService.Save(_config, _currentConfigPath);
            StatusMessage = $"Saved to {_currentConfigPath}";
            SnapshotSaved();
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Save failed", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SaveAs()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "JSON config (*.json)|*.json|All files (*.*)|*.*",
            FileName = Path.GetFileName(_currentConfigPath),
            InitialDirectory = Path.GetDirectoryName(_currentConfigPath)
        };
        if (dlg.ShowDialog() == true)
        {
            CurrentConfigPath = dlg.FileName;
            Save();
        }
    }

    private void Open()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "JSON config (*.json)|*.json|All files (*.*)|*.*",
            InitialDirectory = Path.GetDirectoryName(_currentConfigPath)
        };
        if (dlg.ShowDialog() == true)
        {
            try
            {
                var cfg = ConfigService.Load(dlg.FileName);
                // Preserve the shared contacts list across the swap.
                cfg.Contacts = _config.Contacts;
                MigrateLegacyContacts(cfg);
                StampRecipientStates(cfg);
                foreach (var e in cfg.Emails) EnsureTrailingBlankRecipient(e);
                Config = cfg;
                CurrentConfigPath = dlg.FileName;
                SelectedEmail = cfg.Emails.FirstOrDefault();
                StatusMessage = $"Loaded {dlg.FileName}";
                SnapshotSaved();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Open failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void Preview()
    {
        if (SelectedEmail == null) return;
        var subject = TemplateEngine.Render(SelectedEmail.Subject, SelectedEmail, _config);
        var body = TemplateEngine.Render(SelectedEmail.Body, SelectedEmail, _config);
        var to = string.Join("; ", SelectedEmail.Recipients.Where(r => !r.IsCc).Select(r => r.Email));
        var cc = string.Join("; ", SelectedEmail.Recipients.Where(r =>  r.IsCc).Select(r => r.Email));
        var files = string.Join(Environment.NewLine,
            SelectedEmail.Attachments.SelectMany(a =>
                a.IsFolder
                    ? new[] { $" - [folder] {a.Path}" }
                        .Concat(a.ResolveFiles().Select(f => "     • " + f))
                    : new[] { " - " + a.Path }));
        var ccLine = string.IsNullOrWhiteSpace(cc) ? "" : $"Cc: {cc}\r\n";
        var msg = $"To: {to}\r\n{ccLine}Subject: {subject}\r\n\r\nAttachments:\r\n{files}\r\n\r\n---\r\n{body}";
        MessageBox.Show(msg, "Preview: " + SelectedEmail.Name, MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void Send()
    {
        if (SelectedEmail == null) return;
        var e = SelectedEmail;

        if (!HasAnyToRecipient(e))
        {
            MessageBox.Show("Add at least one recipient in the To field (uncheck CC on at least one row).",
                "No recipients",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var missing = e.Attachments.Where(a => !a.Exists).Select(a => a.Path).ToList();
        if (missing.Count > 0)
        {
            var list = string.Join(Environment.NewLine, missing);
            if (MessageBox.Show("These attachments are missing:\r\n" + list + "\r\n\r\nContinue without them?",
                    "Missing files", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;
        }

        var emptyFolders = e.Attachments
            .Where(a => a.IsFolder && !a.ResolveFiles().Any())
            .Select(a => a.Path).ToList();
        if (emptyFolders.Count > 0)
        {
            var list = string.Join(Environment.NewLine, emptyFolders);
            if (MessageBox.Show("These folder attachments contain no files:\r\n" + list + "\r\n\r\nSend anyway?",
                    "Empty folders", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;
        }

        var mode = _config.SendMode;

        if (mode == SendMode.SendAutomatically)
        {
            var confirm = MessageBox.Show(
                $"Send email '{e.Name}' to {e.Recipients.Count} recipient(s) without review?",
                "Confirm auto-send", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;
        }

        try
        {
            SendEmailCore(e, mode);
            StatusMessage = mode == SendMode.SendAutomatically
                ? $"Sent '{e.Name}' to {e.Recipients.Count} recipient(s)."
                : $"Opened draft for '{e.Name}' in Outlook.";
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "Outlook error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SendAll()
    {
        if (_config.Emails.Count == 0) return;

        var eligible = _config.Emails.Where(HasAnyToRecipient).ToList();
        var skipped  = _config.Emails.Except(eligible).ToList();

        if (eligible.Count == 0)
        {
            MessageBox.Show("No emails have a To recipient. Add at least one row (not CC-only) to an email.",
                "Nothing to send", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var missingByEmail = eligible
            .Select(e => (e, miss: e.Attachments.Where(a => !a.Exists).Select(a => a.Path).ToList()))
            .Where(x => x.miss.Count > 0)
            .ToList();

        var mode = _config.SendMode;
        var verb = mode == SendMode.SendAutomatically ? "Send" : "Open drafts for";
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"{verb} {eligible.Count} email(s):");
        foreach (var e in eligible)
            sb.AppendLine($"  • {e.Name}  ({e.Recipients.Count(r => !r.IsCc)} To, {e.Recipients.Count(r => r.IsCc)} Cc, {e.Attachments.Count} files)");
        if (skipped.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"Skipping {skipped.Count} email(s) with no To recipient:");
            foreach (var e in skipped) sb.AppendLine($"  • {e.Name}");
        }
        if (missingByEmail.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine("Some attachments are missing and will be omitted:");
            foreach (var (e, miss) in missingByEmail)
            {
                sb.AppendLine($"  • {e.Name}:");
                foreach (var p in miss) sb.AppendLine($"      - {p}");
            }
        }
        sb.AppendLine();
        sb.AppendLine("Continue?");

        var icon = mode == SendMode.SendAutomatically ? MessageBoxImage.Warning : MessageBoxImage.Question;
        if (MessageBox.Show(sb.ToString(), "Send all", MessageBoxButton.YesNo, icon) != MessageBoxResult.Yes)
            return;

        int ok = 0;
        var failures = new System.Collections.Generic.List<string>();
        foreach (var e in eligible)
        {
            try
            {
                SendEmailCore(e, mode);
                ok++;
            }
            catch (Exception ex)
            {
                failures.Add($"{e.Name}: {ex.Message}");
            }
        }

        StatusMessage = mode == SendMode.SendAutomatically
            ? $"Sent {ok} of {eligible.Count} email(s)."
            : $"Opened {ok} of {eligible.Count} draft(s) in Outlook.";

        if (failures.Count > 0)
        {
            MessageBox.Show(
                $"{failures.Count} email(s) failed:\r\n\r\n{string.Join(Environment.NewLine, failures)}",
                "Send all — errors", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static bool HasAnyToRecipient(EmailTemplate e) =>
        e.Recipients.Any(r => !r.IsCc && !string.IsNullOrWhiteSpace(r.Email));

    /// <summary>
    /// Renders + sends/drafts a single email. No UI, no validation — caller must check first.
    /// Throws on Outlook errors.
    /// </summary>
    private void SendEmailCore(EmailTemplate e, SendMode mode)
    {
        var subject = TemplateEngine.Render(e.Subject, e, _config);
        var body = TemplateEngine.Render(e.Body, e, _config);
        var toEmails = e.Recipients.Where(r => !r.IsCc && !string.IsNullOrWhiteSpace(r.Email)).Select(r => r.Email);
        var ccEmails = e.Recipients.Where(r =>  r.IsCc && !string.IsNullOrWhiteSpace(r.Email)).Select(r => r.Email);
        var paths = e.Attachments
            .Where(a => a.Exists)
            .SelectMany(a => a.ResolveFiles())
            .Distinct(StringComparer.OrdinalIgnoreCase);

        OutlookService.CreateMail(
            toEmails, ccEmails, subject, body, paths, _config.HtmlBody,
            mode == SendMode.SendAutomatically ? OutlookService.Mode.Send : OutlookService.Mode.Draft);
    }
}
